import os
import sys
import subprocess
import threading
import csv
import datetime
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from dataclasses import dataclass, field

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# 1.4.2 : envoi mail via PowerShell + Outlook
# - Fonctionne avec Outlook Win32 (2016, 2019, 2021, 365)
# - Message propre si Outlook absent ou non compatible (nouveau Outlook)
APP_VERSION = "1.4.2"


@dataclass
class Node:
    path: Path
    name: str
    is_dir: bool
    size: int = 0
    children: list = field(default_factory=list)
    level: int = 0
    access_denied: bool = False  # vrai si dossier non lisible


def human_size(num_bytes: int) -> str:
    """Convertit un nombre d'octets en format lisible."""
    for unit in ["o", "Ko", "Mo", "Go", "To"]:
        if num_bytes < 1024 or unit == "To":
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} To"


def count_entries(root_path: Path) -> int:
    """
    Compte le nombre approximatif d'entrées (fichiers + dossiers)
    pour pouvoir calculer un pourcentage de progression.
    Ignore les erreurs de droits / accès.
    """
    total = 0

    def _onerror(e):
        # On ignore les erreurs (PermissionError, etc.) pendant le comptage
        pass

    for _root, dirs, files in os.walk(root_path, onerror=_onerror):
        total += len(dirs) + len(files)
    return max(total, 1)


def scan_directory(root_path: Path, progress_callback=None):
    """
    Scan récursif du dossier.
    Retourne (root_node, stats_extensions)
    stats_extensions = {".txt": taille_totale, ...}
    """
    ext_stats = {}

    def _scan(path: Path, level: int):
        if path.is_dir():
            node = Node(path=path, name=path.name or str(path), is_dir=True, size=0, level=level)
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        # Une entrée = 1 "élément" pour la progression
                        if progress_callback:
                            progress_callback()

                        child_path = Path(entry.path)
                        try:
                            child_node = _scan(child_path, level + 1)
                            node.children.append(child_node)
                            node.size += child_node.size
                        except (PermissionError, FileNotFoundError, OSError):
                            # On ignore ce qu'on ne peut pas lire
                            continue
            except (PermissionError, FileNotFoundError, OSError):
                # Accès totalement refusé à ce dossier
                node.access_denied = True
            return node
        else:
            # Pour les fichiers, on ne touche plus à la progression :
            # ils sont déjà comptés comme entrées du dossier parent.
            try:
                size = path.stat().st_size
            except (PermissionError, FileNotFoundError, OSError):
                size = 0
            node = Node(path=path, name=path.name, is_dir=False, size=size, level=level)

            ext = path.suffix.lower()
            if not ext:
                ext = "<sans extension>"
            ext_stats[ext] = ext_stats.get(ext, 0) + size

            return node

    root_node = _scan(root_path, 0)
    return root_node, ext_stats


def open_file_in_default_app(path: Path):
    """Ouvre le fichier ou le dossier donné avec l'application par défaut du système."""
    try:
        if os.name == "nt":  # Windows
            os.startfile(str(path))
        elif sys.platform == "darwin":  # macOS
            subprocess.Popen(["open", str(path)])
        else:  # Linux / autres
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


class WinDirScopeApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title(f"WinDirScope v{APP_VERSION} - Analyse d'espace disque")
        self.master.geometry("1000x650")

        self.root_node = None
        self.ext_stats = {}
        self.top_files = []  # Top 100 fichiers les plus volumineux

        self.scan_thread = None
        self.scan_running = False
        self.current_scan_path = None

        self.id_counter = 0
        self.id_to_node = {}

        # Progression
        self.progress_total = 0
        self.progress_current = 0
        self.progress_var = tk.DoubleVar(value=0.0)

        # Profondeur max d'affichage
        self.max_level_var = tk.IntVar(value=5)

        # Contexte menu arbre
        self._context_item_id = None
        self._context_node = None

        # Contexte menu Top 100
        self._top_context_id = None
        self.top_id_to_path = {}

        self._build_ui()

    # ---------- UI ----------

    def _build_ui(self):
        self._build_menu()

        # Barre supérieure
        top_frame = ttk.Frame(self.master)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.btn_select = ttk.Button(
            top_frame, text="Analyser un dossier…", command=self.on_select_folder
        )
        self.btn_select.pack(side=tk.LEFT)

        self.btn_export = ttk.Button(
            top_frame, text="Exporter…", command=self.on_export_results, state="disabled"
        )
        self.btn_export.pack(side=tk.LEFT, padx=(5, 0))

        self.lbl_status = ttk.Label(top_frame, text="Aucun dossier analysé.")
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        self.lbl_level = ttk.Label(top_frame, text="Profondeur analysée (niveaux) :")
        self.lbl_level.pack(side=tk.LEFT, padx=(10, 2))

        self.spin_level = ttk.Spinbox(
            top_frame,
            from_=1,
            to=20,
            textvariable=self.max_level_var,
            width=3,
            command=self.on_change_max_level,
        )
        self.spin_level.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(
            top_frame,
            orient="horizontal",
            mode="determinate",
            variable=self.progress_var,
            maximum=100,
        )
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        # Split vertical : 3 zones
        paned = ttk.PanedWindow(self.master, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Arborescence ---
        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3)

        columns = ("level", "size", "percent")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        self.tree.heading("#0", text="Dossier / Fichier")
        self.tree.heading("level", text="Niveau")
        self.tree.heading("size", text="Taille")
        self.tree.heading("percent", text="Pourcentage")

        self.tree.column("#0", width=450, anchor=tk.W)
        self.tree.column("level", width=70, anchor=tk.CENTER)
        self.tree.column("size", width=120, anchor=tk.E)
        self.tree.column("percent", width=100, anchor=tk.E)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("denied", foreground="red")

        # Menu contextuel arbre
        self.tree_menu = tk.Menu(self.master, tearoff=False)
        self.tree_menu.add_command(label="Ouvrir la cible", command=self.cmd_open_target)
        self.tree_menu.add_command(
            label="Ouvrir le dossier contenant", command=self.cmd_open_containing
        )
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Renommer…", command=self.cmd_rename_node)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(
            label="Supprimer définitivement…", command=self.cmd_delete_node
        )

        self.tree.bind("<Button-3>", self.on_tree_right_click)

        # --- Répartition par extension ---
        ext_frame = ttk.LabelFrame(paned, text="Répartition par extension")
        paned.add(ext_frame, weight=2)

        ext_columns = ("ext", "size", "percent")
        self.ext_tree = ttk.Treeview(
            ext_frame, columns=ext_columns, show="headings", height=4
        )
        self.ext_tree.heading("ext", text="Extension")
        self.ext_tree.heading("size", text="Taille totale")
        self.ext_tree.heading("percent", text="Pourcentage")

        self.ext_tree.column("ext", width=120, anchor=tk.W)
        self.ext_tree.column("size", width=120, anchor=tk.E)
        self.ext_tree.column("percent", width=100, anchor=tk.E)

        vsb_ext = ttk.Scrollbar(ext_frame, orient="vertical", command=self.ext_tree.yview)
        self.ext_tree.configure(yscrollcommand=vsb_ext.set)

        self.ext_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb_ext.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Top 100 fichiers ---
        top_frame = ttk.LabelFrame(paned, text="Top 100 fichiers les plus volumineux")
        paned.add(top_frame, weight=3)

        top_columns = ("size", "percent", "path")
        self.top_tree = ttk.Treeview(
            top_frame, columns=top_columns, show="headings", height=6
        )
        self.top_tree.heading("size", text="Taille")
        self.top_tree.heading("percent", text="% du total")
        self.top_tree.heading("path", text="Chemin complet")

        self.top_tree.column("size", width=120, anchor=tk.E)
        self.top_tree.column("percent", width=100, anchor=tk.E)
        self.top_tree.column("path", width=500, anchor=tk.W)

        vsb_top = ttk.Scrollbar(top_frame, orient="vertical", command=self.top_tree.yview)
        self.top_tree.configure(yscrollcommand=vsb_top.set)

        self.top_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb_top.pack(side=tk.RIGHT, fill=tk.Y)

        # Menu contextuel Top 100
        self.top_tree_menu = tk.Menu(self.master, tearoff=False)
        self.top_tree_menu.add_command(
            label="Ouvrir la cible", command=self.cmd_top_open_target
        )
        self.top_tree_menu.add_command(
            label="Ouvrir le dossier contenant", command=self.cmd_top_open_containing
        )
        self.top_tree_menu.add_separator()
        the_cmd = self.cmd_top_rename
        self.top_tree_menu.add_command(label="Renommer…", command=the_cmd)
        self.top_tree_menu.add_separator()
        self.top_tree_menu.add_command(
            label="Supprimer définitivement…", command=self.cmd_top_delete
        )

        self.top_tree.bind("<Button-3>", self.on_top_right_click)

    def _build_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Analyser un dossier…", command=self.on_select_folder)
        file_menu.add_command(
            label="Exporter les résultats…", command=self.on_export_results
        )

        # Sous-menu "Envoyer le rapport par mail"
        send_menu = tk.Menu(file_menu, tearoff=False)
        send_menu.add_command(label="HTML", command=lambda: self.on_send_report("html"))
        send_menu.add_command(label="CSV", command=lambda: self.on_send_report("csv"))
        send_menu.add_command(label="JSON", command=lambda: self.on_send_report("json"))
        send_menu.add_command(label="TXT", command=lambda: self.on_send_report("txt"))
        file_menu.add_cascade(label="Envoyer le rapport par mail", menu=send_menu)

        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.master.quit)
        menubar.add_cascade(label="Fichier", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="À propos", command=self.on_about)
        menubar.add_cascade(label="Aide", menu=help_menu)

    # ---------- Callbacks ----------

    def on_about(self):
        messagebox.showinfo(
            "À propos",
            (
                "WinDirScope - Analyse d'espace disque\n"
                f"Version : {APP_VERSION}\n"
                "Auteur : Vincent TOUZOT\n"
                "Outil développé avec l'assistance de ChatGPT "
                "(modèle GPT-5.1 Thinking, OpenAI).\n"
                "\n"
                "Fonctions principales :\n"
                "- Analyse de dossiers locaux et réseaux\n"
                "- Affichage arborescent avec tailles et pourcentages\n"
                "- Répartition par extension\n"
                "- Top 100 fichiers les plus volumineux\n"
                "- Menus clic droit (ouvrir, renommer, supprimer)\n"
                "- Export CSV / JSON / TXT / HTML (rapport HTML interactif)\n"
                "- Envoi du rapport par mail (Outlook via PowerShell)\n"
                "- Indication visuelle des dossiers en accès refusé.\n"
            ),
        )

    def on_select_folder(self):
        if self.scan_running:
            messagebox.showwarning("Analyse en cours", "Une analyse est déjà en cours.")
            return

        folder = filedialog.askdirectory(title="Choisir un dossier à analyser")
        if not folder:
            return

        path = Path(folder)
        if not path.exists():
            messagebox.showerror(
                "Erreur", "Ce dossier n'existe pas ou n'est pas accessible."
            )
            return

        self.current_scan_path = path
        self.progress_total = count_entries(path)
        self.progress_current = 0
        self.progress_var.set(0.0)

        self.scan_running = True
        self.btn_export.config(state="disabled")
        self.lbl_status.config(text=f"Analyse en cours : {folder}")
        self._clear_views()

        self.scan_thread = threading.Thread(
            target=self._scan_worker, args=(path,), daemon=True
        )
        self.scan_thread.start()
        self.master.after(200, self._poll_scan_thread)

    def _scan_worker(self, path: Path):
        try:
            root_node, ext_stats = scan_directory(path, progress_callback=self._progress_tick)
            self.root_node = root_node
            self.ext_stats = ext_stats
            self._scan_error = None
        except Exception as e:
            self.root_node = None
            self.ext_stats = {}
            self._scan_error = e

    def _progress_tick(self):
        self.progress_current += 1

    def _poll_scan_thread(self):
        if self.scan_thread is None:
            return

        self._update_progress_ui()

        if self.scan_thread.is_alive():
            self.master.after(200, self._poll_scan_thread)
        else:
            self.scan_running = False
            if getattr(self, "_scan_error", None):
                messagebox.showerror("Erreur d'analyse", str(self._scan_error))
                self.lbl_status.config(text="Erreur lors de l'analyse.")
            else:
                self.lbl_status.config(text=f"Analyse terminée : {self.root_node.path}")
                self._compute_top_files()
                self._populate_views()
                self.btn_export.config(state="normal")

    def _update_progress_ui(self):
        if self.progress_total <= 0:
            self.progress_var.set(0.0)
            return
        percent = min(100.0, (self.progress_current / self.progress_total) * 100.0)
        self.progress_var.set(percent)
        if self.current_scan_path:
            self.lbl_status.config(
                text=(
                    f"Analyse en cours : {self.current_scan_path} "
                    f"({self.progress_current}/{self.progress_total} éléments, {percent:5.1f} %)"
                )
            )

    def on_change_max_level(self):
        if not self.root_node or self.scan_running:
            return
        self._clear_tree_view()
        self._populate_tree_view()

    # ---------- Affichage ----------

    def _clear_views(self):
        self._clear_tree_view()
        self.ext_tree.delete(*self.ext_tree.get_children())
        self.top_tree.delete(*self.top_tree.get_children())
        self.id_to_node.clear()
        self.id_counter = 0
        self.top_files = []
        self.top_id_to_path = {}

    def _clear_tree_view(self):
        self.tree.delete(*self.tree.get_children())
        self.id_to_node.clear()
        self.id_counter = 0

    def _next_id(self):
        self.id_counter += 1
        return f"node_{self.id_counter}"

    def _populate_views(self):
        if not self.root_node:
            return
        self._populate_tree_view()
        self._populate_ext_view()
        self._populate_top_files_view()

    def _populate_tree_view(self):
        if not self.root_node:
            return

        total_size = self.root_node.size or 1

        try:
            max_level = int(self.max_level_var.get())
        except (TypeError, ValueError):
            max_level = 5

        def add_node(parent_id, node: Node):
            if node.level > max_level:
                return

            text = node.name
            tags = ()
            if node.access_denied:
                text = f"{node.name} [ACCÈS REFUSÉ]"
                tags = ("denied",)

            tree_id = self._next_id()
            self.id_to_node[tree_id] = node
            percent = (node.size / total_size) * 100
            self.tree.insert(
                parent_id,
                "end",
                iid=tree_id,
                text=text,
                values=(node.level, human_size(node.size), f"{percent:5.2f} %"),
                tags=tags,
            )
            if node.is_dir:
                for child in sorted(node.children, key=lambda n: n.size, reverse=True):
                    add_node(tree_id, child)

        add_node("", self.root_node)
        first = self.tree.get_children()
        if first:
            self.tree.item(first[0], open=True)

    def _populate_ext_view(self):
        self.ext_tree.delete(*self.ext_tree.get_children())
        total_ext_size = sum(self.ext_stats.values()) or 1
        for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
            percent = (size / total_ext_size) * 100
            self.ext_tree.insert(
                "",
                "end",
                values=(ext, human_size(size), f"{percent:5.2f} %"),
            )

    def _populate_top_files_view(self):
        self.top_tree.delete(*self.top_tree.get_children())
        self.top_id_to_path = {}
        for row in self.top_files:
            item_id = self.top_tree.insert(
                "",
                "end",
                values=(
                    row["size_human"],
                    f"{row['percent_total']:.2f} %",
                    row["path"],
                ),
            )
            self.top_id_to_path[item_id] = Path(row["path"])

    # ---------- Top 100 ----------

    def _compute_top_files(self):
        self.top_files = []
        if not self.root_node:
            return

        total_size = self.root_node.size or 1
        files = []

        def visit(node: Node):
            if node.is_dir:
                for c in node.children:
                    visit(c)
            else:
                percent = (node.size / total_size) * 100
                files.append(
                    {
                        "path": str(node.path),
                        "name": node.name,
                        "size_bytes": node.size,
                        "size_human": human_size(node.size),
                        "percent_total": percent,
                        "level": node.level,
                    }
                )

        visit(self.root_node)
        files.sort(key=lambda r: r["size_bytes"], reverse=True)
        self.top_files = files[:100]

    # ---------- Menu contextuel arbre ----------

    def on_tree_right_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.tree.selection_set(item_id)
        node = self.id_to_node.get(item_id)
        if not node:
            return
        self._context_item_id = item_id
        self._context_node = node
        try:
            self.tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_menu.grab_release()

    def _get_context_node(self):
        if not self._context_item_id:
            return None, None
        return self._context_item_id, self._context_node

    def cmd_open_target(self):
        item_id, node = self._get_context_node()
        if not node:
            return
        path = node.path
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return
        open_file_in_default_app(path)

    def cmd_open_containing(self):
        item_id, node = self._get_context_node()
        if not node:
            return
        path = node.path
        folder = path if node.is_dir else path.parent
        if not folder.exists():
            messagebox.showerror("Chemin introuvable", f"Le dossier n'existe plus :\n{folder}")
            return
        open_file_in_default_app(folder)

    def cmd_rename_node(self):
        item_id, node = self._get_context_node()
        if not node:
            return

        if node is self.root_node:
            messagebox.showwarning(
                "Renommage refusé",
                "Le dossier racine ne peut pas être renommé ici.",
            )
            return

        old_name = node.name
        new_name = simpledialog.askstring(
            "Renommer",
            f"Nouveau nom pour :\n{node.path}",
            initialvalue=old_name,
            parent=self.master,
        )
        if not new_name or new_name == old_name:
            return

        old_path = node.path
        new_path = old_path.with_name(new_name)

        if new_path.exists():
            messagebox.showerror(
                "Existe déjà",
                f"Un élément portant ce nom existe déjà :\n{new_path}",
            )
            return

        try:
            os.rename(old_path, new_path)
        except Exception as e:
            messagebox.showerror("Erreur de renommage", f"Impossible de renommer :\n{e}")
            return

        node.path = new_path
        node.name = new_name

        if node.is_dir:
            self._update_child_paths(node, old_path, new_path)

        self.tree.item(item_id, text=new_name)

        messagebox.showinfo(
            "Renommage effectué",
            "Le renommage a été effectué sur le système de fichiers.\n"
            "Les statistiques affichées ne sont pas recalculées automatiquement.\n"
            "Relancez une analyse si nécessaire.",
        )

    def _update_child_paths(self, parent_node: Node, old_root: Path, new_root: Path):
        for child in parent_node.children:
            try:
                rel = child.path.relative_to(old_root)
                child.path = new_root / rel
            except ValueError:
                pass
            if child.is_dir:
                self._update_child_paths(child, old_root, new_root)

    def cmd_delete_node(self):
        item_id, node = self._get_context_node()
        if not node:
            return

        if node is self.root_node:
            messagebox.showwarning(
                "Suppression refusée",
                "Le dossier racine ne peut pas être supprimé depuis cette interface.",
            )
            return

        path = node.path
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return

        type_txt = "dossier" if node.is_dir else "fichier"
        if not messagebox.askyesno(
            "Suppression définitive",
            f"ATTENTION : ceci va SUPPRIMER DÉFINITIVEMENT le {type_txt} :\n\n"
            f"{path}\n\n"
            "Cette action ne passe pas par la corbeille et est irréversible.\n\n"
            "Confirmer la suppression ?",
            icon=messagebox.WARNING,
        ):
            return

        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                os.remove(path)
        except PermissionError as e:
            messagebox.showerror(
                "Erreur de suppression",
                "Impossible de supprimer (accès refusé).\n\n"
                "Causes possibles :\n"
                "- fichier/dossier ouvert dans une autre application,\n"
                "- antivirus ou logiciel de sauvegarde qui le verrouille,\n"
                "- droits NTFS insuffisants.\n\n"
                f"Détail :\n{e}",
            )
            return
        except Exception as e:
            messagebox.showerror("Erreur de suppression", f"Impossible de supprimer :\n{e}")
            return

        parent_item_id = self.tree.parent(item_id)
        parent_node = self.id_to_node.get(parent_item_id)

        if parent_node:
            parent_node.children = [c for c in parent_node.children if c is not node]

        if item_id in self.id_to_node:
            del self.id_to_node[item_id]

        self.tree.delete(item_id)

        messagebox.showinfo(
            "Suppression effectuée",
            "Le fichier / dossier a été supprimé du système de fichiers.\n"
            "Les statistiques affichées ne sont pas recalculées automatiquement.\n"
            "Relancez une analyse si nécessaire.",
        )

    # ---------- Menu contextuel Top 100 ----------

    def on_top_right_click(self, event):
        item_id = self.top_tree.identify_row(event.y)
        if not item_id:
            return
        self.top_tree.selection_set(item_id)
        self._top_context_id = item_id
        try:
            self.top_tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.top_tree_menu.grab_release()

    def _get_top_context_path(self):
        if not self._top_context_id:
            return None
        return self.top_id_to_path.get(self._top_context_id)

    def cmd_top_open_target(self):
        path = self._get_top_context_path()
        if not path:
            return
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return
        open_file_in_default_app(path)

    def cmd_top_open_containing(self):
        path = self._get_top_context_path()
        if not path:
            return
        folder = path if path.is_dir() else path.parent
        if not folder.exists():
            messagebox.showerror("Chemin introuvable", f"Le dossier n'existe plus :\n{folder}")
            return
        open_file_in_default_app(folder)

    def cmd_top_rename(self):
        path = self._get_top_context_path()
        if not path:
            return
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return

        old_name = path.name
        new_name = simpledialog.askstring(
            "Renommer",
            f"Nouveau nom pour :\n{path}",
            initialvalue=old_name,
            parent=self.master,
        )
        if not new_name or new_name == old_name:
            return

        new_path = path.with_name(new_name)
        if new_path.exists():
            messagebox.showerror("Existe déjà", f"Un élément portant ce nom existe déjà :\n{new_path}")
            return

        try:
            os.rename(path, new_path)
        except Exception as e:
            messagebox.showerror("Erreur de renommage", f"Impossible de renommer :\n{e}")
            return

        self.top_id_to_path[self._top_context_id] = new_path

        for r in self.top_files:
            if Path(r["path"]) == path:
                r["path"] = str(new_path)
                r["name"] = new_name
                break

        values = list(self.top_tree.item(self._top_context_id, "values"))
        if len(values) >= 3:
            values[2] = str(new_path)
            self.top_tree.item(self._top_context_id, values=values)

        messagebox.showinfo(
            "Renommage effectué",
            "Le renommage a été effectué sur le système de fichiers.\n"
            "Les statistiques affichées ne sont pas recalculées automatiquement.\n"
            "Relancez une analyse si nécessaire.",
        )

    def cmd_top_delete(self):
        path = self._get_top_context_path()
        if not path:
            return
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return

        type_txt = "dossier" if path.is_dir() else "fichier"
        if not messagebox.askyesno(
            "Suppression définitive",
            f"ATTENTION : ceci va SUPPRIMER DÉFINITIVEMENT le {type_txt} :\n\n"
            f"{path}\n\n"
            "Cette action ne passe pas par la corbeille et est irréversible.\n\n"
            "Confirmer la suppression ?",
            icon=messagebox.WARNING,
        ):
            return

        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                os.remove(path)
        except PermissionError as e:
            messagebox.showerror(
                "Erreur de suppression",
                "Impossible de supprimer (accès refusé).\n\n"
                "Causes possibles :\n"
                "- fichier/dossier ouvert dans une autre application,\n"
                "- antivirus ou logiciel de sauvegarde qui le verrouille,\n"
                "- droits NTFS insuffisants.\n\n"
                f"Détail :\n{e}",
            )
            return
        except Exception as e:
            messagebox.showerror("Erreur de suppression", f"Impossible de supprimer :\n{e}")
            return

        if self._top_context_id in self.top_id_to_path:
            del self.top_id_to_path[self._top_context_id]
        self.top_tree.delete(self._top_context_id)

        self.top_files = [r for r in self.top_files if Path(r["path"]) != path]

        messagebox.showinfo(
            "Suppression effectuée",
            "Le fichier / dossier a été supprimé du système de fichiers.\n"
            "Les statistiques affichées ne sont pas recalculées automatiquement.\n"
            "Relancez une analyse si nécessaire.",
        )

    # ---------- Flatten pour export ----------

    def _flatten_tree(self):
        rows = []
        total_size = self.root_node.size or 1

        def visit(node: Node):
            percent = (node.size / total_size) * 100
            rows.append(
                {
                    "path": str(node.path),
                    "name": node.name,
                    "level": node.level,
                    "type": "dossier" if node.is_dir else "fichier",
                    "size_bytes": node.size,
                    "size_human": human_size(node.size),
                    "percent_total": percent,
                    "access_denied": node.access_denied,
                }
            )
            for child in node.children:
                visit(child)

        visit(self.root_node)
        return rows

    # ---------- Envoi par mail ----------

    def on_send_report(self, fmt: str):
        if not self.root_node:
            messagebox.showwarning(
                "Aucun résultat",
                "Aucun dossier n'a encore été analysé. Lancez une analyse avant d'envoyer un rapport.",
            )
            return

        self._compute_top_files()

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        raw_root_name = self.root_node.name or "analyse"
        invalid_chars = '<>:"/\\|?*'
        root_name = "".join(c for c in raw_root_name if c not in invalid_chars)
        if not root_name.strip():
            root_name = "racine"

        tmp_dir = Path(tempfile.gettempdir())

        try:
            if fmt == "html":
                filename = f"WinDirScope_{timestamp}_{root_name}.html"
                html_path = tmp_dir / filename
                self._export_html(html_path)
                attach_path = html_path
            elif fmt == "csv":
                base = f"WinDirScope_{timestamp}_{root_name}_csv"
                tree_file = tmp_dir / f"{base}_arborescence.csv"
                ext_file = tmp_dir / f"{base}_extensions.csv"
                top_file = tmp_dir / f"{base}_top100.csv"
                self._export_tree_csv(tree_file)
                self._export_ext_csv(ext_file)
                self._export_top_csv(top_file)
                zip_path = tmp_dir / f"{base}.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(tree_file, tree_file.name)
                    z.write(ext_file, ext_file.name)
                    z.write(top_file, top_file.name)
                attach_path = zip_path
            elif fmt == "json":
                base = f"WinDirScope_{timestamp}_{root_name}_json"
                tree_file = tmp_dir / f"{base}_arborescence.json"
                ext_file = tmp_dir / f"{base}_extensions.json"
                top_file = tmp_dir / f"{base}_top100.json"
                self._export_tree_json(tree_file)
                self._export_ext_json(ext_file)
                self._export_top_json(top_file)
                zip_path = tmp_dir / f"{base}.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(tree_file, tree_file.name)
                    z.write(ext_file, ext_file.name)
                    z.write(top_file, top_file.name)
                attach_path = zip_path
            elif fmt == "txt":
                base = f"WinDirScope_{timestamp}_{root_name}_txt"
                tree_file = tmp_dir / f"{base}_arborescence.txt"
                ext_file = tmp_dir / f"{base}_extensions.txt"
                top_file = tmp_dir / f"{base}_top100.txt"
                self._export_tree_txt(tree_file)
                self._export_ext_txt(ext_file)
                self._export_top_txt(top_file)
                zip_path = tmp_dir / f"{base}.zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(tree_file, tree_file.name)
                    z.write(ext_file, ext_file.name)
                    z.write(top_file, top_file.name)
                attach_path = zip_path
            else:
                messagebox.showerror("Format inconnu", f"Format non géré : {fmt}")
                return
        except Exception as e:
            messagebox.showerror(
                "Erreur de génération", f"Impossible de générer le rapport :\n{e}"
            )
            return

        self._open_email_with_attachment(attach_path)

    def _open_email_with_attachment(self, filepath: Path):
        """
        Tente d'ouvrir un nouveau mail avec la pièce jointe.
        - Sous Windows : via PowerShell + COM Outlook (Outlook 2016+ Win32).
        - Si Outlook absent ou non compatible (nouveau Outlook, etc.) :
          message clair, sans erreur technique.
        """
        if os.name != "nt":
            messagebox.showinfo(
                "Fichier prêt",
                f"Le rapport a été généré :\n{filepath}\n\n"
                "L'ouverture automatique du client mail n'est disponible que sous Windows.\n"
                "Créez un nouveau message et ajoutez ce fichier en pièce jointe.",
            )
            return

        subject = f"Rapport WinDirScope - {self.root_node.name}"
        body = (
            "Bonjour,`n`n"
            "Veuillez trouver en pièce jointe le rapport généré par WinDirScope.`n`n"
            "Cordialement,`n"
            "WinDirScope"
        )

        # Échappement simple des apostrophes pour PowerShell
        ps_path = str(filepath).replace("'", "''")
        ps_subject = subject.replace("'", "''")
        ps_body = body.replace("'", "''")

        ps_script = (
            "$ErrorActionPreference = 'Stop'; "
            "try { "
            "  $ol = New-Object -ComObject Outlook.Application; "
            "  if (-not $ol) { throw 'Outlook non disponible'; } "
            "  $mail = $ol.CreateItem(0); "
            f"  $mail.Subject = '{ps_subject}'; "
            f"  $mail.Body = '{ps_body}'; "
            f"  $null = $mail.Attachments.Add('{ps_path}'); "
            "  $mail.Display(); "
            "  exit 0; "
            "} catch { "
            "  Write-Error $_.Exception.Message; "
            "  exit 1; "
            "}"
        )

        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_script],
                capture_output=True,
                text=True,
            )
        except Exception:
            result = None

        if not result or result.returncode != 0:
            # Outlook non disponible (non installé ou version non compatible COM, ex : "nouveau Outlook")
            messagebox.showinfo(
                "Outlook indisponible",
                f"Le rapport a été généré :\n{filepath}\n\n"
                "Outlook (version classique 2016/2019/2021/365) ne semble pas disponible "
                "ou ne permet pas l'automatisation.\n"
                "Créez un nouveau message dans votre messagerie et ajoutez ce fichier en pièce jointe.",
            )

    # ---------- Export classique ----------

    def on_export_results(self):
        if not self.root_node:
            messagebox.showwarning(
                "Aucun résultat",
                "Aucun dossier n'a encore été analysé. Lancez une analyse avant d'exporter.",
            )
            return

        self._compute_top_files()

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        raw_root_name = self.root_node.name or "analyse"
        invalid_chars = '<>:"/\\|?*'
        root_name = "".join(c for c in raw_root_name if c not in invalid_chars)
        if not root_name.strip():
            root_name = "racine"

        default_name = f"WinDirScope_{timestamp}_{root_name}.html"

        base_path_str = filedialog.asksaveasfilename(
            title="Exporter les résultats (nom de base du fichier)",
            defaultextension=".html",
            filetypes=[
                ("Page HTML", "*.html"),
                ("Page HTML (htm)", "*.htm"),
                ("Fichier CSV", "*.csv"),
                ("Fichier JSON", "*.json"),
                ("Fichier texte", "*.txt"),
            ],
            initialfile=default_name,
        )
        if not base_path_str:
            return

        base_path = Path(base_path_str)
        suffix = base_path.suffix.lower()

        exported_files = []

        try:
            if suffix == ".json":
                tree_file = base_path.with_name(base_path.stem + "_arborescence.json")
                ext_file = base_path.with_name(base_path.stem + "_extensions.json")
                top_file = base_path.with_name(base_path.stem + "_top100.json")
                self._export_tree_json(tree_file)
                self._export_ext_json(ext_file)
                self._export_top_json(top_file)
                exported_files = [tree_file, ext_file, top_file]
            elif suffix == ".txt":
                tree_file = base_path.with_name(base_path.stem + "_arborescence.txt")
                ext_file = base_path.with_name(base_path.stem + "_extensions.txt")
                top_file = base_path.with_name(base_path.stem + "_top100.txt")
                self._export_tree_txt(tree_file)
                self._export_ext_txt(ext_file)
                self._export_top_txt(top_file)
                exported_files = [tree_file, ext_file, top_file]
            elif suffix in (".html", ".htm"):
                html_file = base_path
                self._export_html(html_file)
                exported_files = [html_file]
            else:  # CSV
                tree_file = base_path.with_name(base_path.stem + "_arborescence.csv")
                ext_file = base_path.with_name(base_path.stem + "_extensions.csv")
                top_file = base_path.with_name(base_path.stem + "_top100.csv")
                self._export_tree_csv(tree_file)
                self._export_ext_csv(ext_file)
                self._export_top_csv(top_file)
                exported_files = [tree_file, ext_file, top_file]
        except Exception as e:
            messagebox.showerror(
                "Erreur d'export", f"Impossible d'exporter les résultats : {e}"
            )
            return

        if exported_files:
            open_file_in_default_app(exported_files[0])

        msg = "Résultats exportés :\n" + "\n".join(f"- {p}" for p in exported_files)
        messagebox.showinfo("Export terminé", msg)

    # --- CSV ---

    def _export_tree_csv(self, filepath: Path):
        rows = self._flatten_tree()
        with filepath.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(
                [
                    "Chemin complet",
                    "Nom",
                    "Niveau",
                    "Type",
                    "Taille (octets)",
                    "Taille lisible",
                    "% du total",
                    "Accès refusé",
                ]
            )
            for row in rows:
                writer.writerow(
                    [
                        row["path"],
                        row["name"],
                        row["level"],
                        row["type"],
                        row["size_bytes"],
                        row["size_human"],
                        f"{row['percent_total']:.4f}",
                        "Oui" if row["access_denied"] else "Non",
                    ]
                )

    def _export_ext_csv(self, filepath: Path):
        total_ext_size = sum(self.ext_stats.values()) or 1
        with filepath.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(
                ["Extension", "Taille totale (octets)", "Taille lisible", "% du total"]
            )
            for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
                percent = (size / total_ext_size) * 100
                writer.writerow(
                    [ext, size, human_size(size), f"{percent:.4f}"]
                )

    def _export_top_csv(self, filepath: Path):
        with filepath.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(
                [
                    "Chemin complet",
                    "Nom",
                    "Taille (octets)",
                    "Taille lisible",
                    "% du total",
                    "Niveau",
                ]
            )
            for row in self.top_files:
                writer.writerow(
                    [
                        row["path"],
                        row["name"],
                        row["size_bytes"],
                        row["size_human"],
                        f"{row['percent_total']:.4f}",
                        row["level"],
                    ]
                )

    # --- JSON ---

    def _export_tree_json(self, filepath: Path):
        rows = self._flatten_tree()
        with filepath.open("w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)

    def _export_ext_json(self, filepath: Path):
        total_ext_size = sum(self.ext_stats.values()) or 1
        items = []
        for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
            percent = (size / total_ext_size) * 100
            items.append(
                {
                    "extension": ext,
                    "size_bytes": size,
                    "size_human": human_size(size),
                    "percent_total": percent,
                }
            )
        with filepath.open("w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)

    def _export_top_json(self, filepath: Path):
        with filepath.open("w", encoding="utf-8") as f:
            json.dump(self.top_files, f, ensure_ascii=False, indent=2)

    # --- TXT ---

    def _export_tree_txt(self, filepath: Path):
        rows = self._flatten_tree()
        with filepath.open("w", encoding="utf-8") as f:
            for row in rows:
                indent = "  " * int(row["level"])
                line = (
                    f"{indent}{row['name']} "
                    f"({row['type']}, {row['size_human']}, "
                    f"{row['percent_total']:.2f} %, "
                    f"accès refusé: {'Oui' if row['access_denied'] else 'Non'}) "
                    f"- {row['path']}"
                )
                f.write(line + "\n")

    def _export_ext_txt(self, filepath: Path):
        total_ext_size = sum(self.ext_stats.values()) or 1
        with filepath.open("w", encoding="utf-8") as f:
            for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
                percent = (size / total_ext_size) * 100
                line = f"{ext}: {human_size(size)} ({percent:.2f} %, {size} octets)"
                f.write(line + "\n")

    def _export_top_txt(self, filepath: Path):
        with filepath.open("w", encoding="utf-8") as f:
            for row in self.top_files:
                line = (
                    f"{row['name']} "
                    f"({row['size_human']}, {row['percent_total']:.2f} %, niveau {row['level']}) "
                    f"- {row['path']}"
                )
                f.write(line + "\n")

    # --- HTML ---

    def _export_html(self, filepath: Path):
        def esc(s: str) -> str:
            return (
                s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
            )

        total_size = self.root_node.size or 1
        total_ext_size = sum(self.ext_stats.values()) or 1

        def node_to_html(node: Node) -> str:
            percent = (node.size / total_size) * 100
            name_raw = node.name
            name = esc(name_raw)
            name_lc = esc(name_raw.lower())
            path = esc(str(node.path))
            size_h = esc(human_size(node.size))
            type_txt = "dossier" if node.is_dir else "fichier"
            lvl = node.level
            denied = node.access_denied

            info = (
                f"{type_txt}, niveau {lvl}, {size_h}, "
                f"{percent:.2f} %, "
                f"{'ACCÈS REFUSÉ' if denied else 'OK'}"
            )

            line = (
                f'<span class="name">{name}</span> '
                f'<span class="meta">({esc(info)})</span><br>'
                f'<span class="path">{path}</span>'
            )

            if node.is_dir:
                open_attr = " open" if node.level <= 1 else ""
                classes = "dir node denied" if denied else "dir node"
                html = [
                    f'<details{open_attr}>'
                    f'<summary class="{classes}" '
                    f'data-name="{name_lc}" data-level="{lvl}" data-type="dir">'
                    f"{line}</summary>",
                ]
                if node.children:
                    html.append("<ul>")
                    for child in sorted(node.children, key=lambda n: n.size, reverse=True):
                        html.append("<li>")
                        html.append(node_to_html(child))
                        html.append("</li>")
                    html.append("</ul>")
                html.append("</details>")
                return "".join(html)
            else:
                classes = "file node denied" if denied else "file node"
                return (
                    f'<div class="{classes}" '
                    f'data-name="{name_lc}" data-level="{lvl}" data-type="file">'
                    f"{line}</div>"
                )

        html_parts = []
        html_parts.append("<!DOCTYPE html>")
        html_parts.append("<html lang='fr'>")
        html_parts.append("<head>")
        html_parts.append("<meta charset='utf-8'>")
        html_parts.append(
            f"<title>WinDirScope - Rapport {esc(self.root_node.name)}</title>"
        )
        html_parts.append("<style>")
        html_parts.append(
            """
            body {
                font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                font-size: 14px;
                background: #f5f5f5;
                color: #222;
                margin: 0;
                padding: 0;
            }
            header {
                background: #2c3e50;
                color: #ecf0f1;
                padding: 10px 16px;
            }
            header h1 {
                margin: 0;
                font-size: 18px;
            }
            header .subtitle {
                font-size: 12px;
                opacity: 0.9;
            }
            main {
                display: grid;
                grid-template-columns: 2fr 1fr 2fr;
                gap: 16px;
                padding: 16px;
            }
            section {
                background: #ffffff;
                border-radius: 6px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.08);
                padding: 12px 16px;
                box-sizing: border-box;
                max-height: 80vh;
                overflow: auto;
            }
            #filters {
                display: flex;
                flex-wrap: wrap;
                gap: 8px;
                align-items: center;
                margin-bottom: 8px;
                font-size: 12px;
            }
            #filters label {
                display: flex;
                flex-direction: column;
                gap: 2px;
            }
            #filters input {
                padding: 2px 4px;
                font-size: 12px;
            }
            #filters button {
                padding: 2px 8px;
                font-size: 12px;
                cursor: pointer;
            }
            details {
                margin-left: 8px;
                margin-top: 4px;
            }
            summary {
                cursor: pointer;
            }
            .name {
                font-weight: 600;
            }
            .meta {
                font-size: 11px;
                color: #555;
            }
            .path {
                font-size: 11px;
                color: #888;
            }
            .dir {
                color: #2c3e50;
            }
            .file {
                margin-left: 20px;
            }
            .denied {
                color: #c0392b;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                font-size: 12px;
            }
            th, td {
                border-bottom: 1px solid #ddd;
                padding: 4px 6px;
                text-align: right;
            }
            th:first-child, td:first-child {
                text-align: left;
            }
            th {
                background: #f0f0f0;
                position: sticky;
                top: 0;
                z-index: 1;
            }
            footer {
                font-size: 11px;
                color: #666;
                padding: 8px 16px 12px 16px;
            }
        """
        )
        html_parts.append("</style>")
        html_parts.append("</head>")
        html_parts.append("<body>")

        html_parts.append("<header>")
        html_parts.append("<h1>WinDirScope - Rapport d'analyse</h1>")
        html_parts.append("<div class='subtitle'>")
        html_parts.append(f"Dossier racine : {esc(str(self.root_node.path))}<br>")
        html_parts.append(f"Taille totale : {esc(human_size(self.root_node.size))}")
        html_parts.append("</div>")
        html_parts.append("</header>")

        html_parts.append("<main>")

        # Arborescence
        html_parts.append("<section class='tree'>")
        html_parts.append("<h2>Arborescence</h2>")
        html_parts.append(
            """
        <div id="filters">
          <label>
            Nom contient :
            <input type="text" id="filter-name" placeholder="texte à rechercher">
          </label>
          <label>
            Niveau max :
            <input type="number" id="filter-level" min="0" placeholder="ex : 3">
          </label>
          <button type="button" id="filter-apply">Appliquer</button>
          <button type="button" id="filter-reset">Réinitialiser</button>
        </div>
        """
        )
        html_parts.append(node_to_html(self.root_node))
        html_parts.append("</section>")

        # Extensions
        html_parts.append("<section class='ext'>")
        html_parts.append("<h2>Répartition par extension</h2>")
        html_parts.append("<table>")
        html_parts.append(
            "<thead><tr><th>Extension</th><th>Taille lisible</th>"
            "<th>Taille (octets)</th><th>% du total</th></tr></thead>"
        )
        html_parts.append("<tbody>")
        for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
            percent = (size / total_ext_size) * 100 if total_ext_size > 0 else 0.0
            html_parts.append(
                "<tr>"
                f"<td>{esc(ext)}</td>"
                f"<td>{esc(human_size(size))}</td>"
                f"<td>{size}</td>"
                f"<td>{percent:.2f}</td>"
                "</tr>"
            )
        html_parts.append("</tbody></table>")
        html_parts.append("</section>")

        # Top 100
        html_parts.append("<section class='top'>")
        html_parts.append("<h2>Top 100 fichiers les plus volumineux</h2>")
        html_parts.append("<table>")
        html_parts.append(
            "<thead><tr><th>Nom</th><th>Taille lisible</th>"
            "<th>Taille (octets)</th><th>% du total</th><th>Chemin complet</th></tr></thead>"
        )
        html_parts.append("<tbody>")
        for row in self.top_files:
            html_parts.append(
                "<tr>"
                f"<td>{esc(row['name'])}</td>"
                f"<td>{esc(row['size_human'])}</td>"
                f"<td>{row['size_bytes']}</td>"
                f"<td>{row['percent_total']:.2f}</td>"
                f"<td>{esc(row['path'])}</td>"
                "</tr>"
            )
        html_parts.append("</tbody></table>")
        html_parts.append("</section>")

        html_parts.append("</main>")

        html_parts.append("<footer>")
        html_parts.append(
            f"Rapport généré par WinDirScope v{APP_VERSION} le "
            f"{datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}."
        )
        html_parts.append("</footer>")

        html_parts.append(
            """
<script>
(function() {
  const nameInput = document.getElementById('filter-name');
  const levelInput = document.getElementById('filter-level');
  const applyBtn = document.getElementById('filter-apply');
  const resetBtn = document.getElementById('filter-reset');

  function applyFilters() {
    const nameFilter = (nameInput && nameInput.value || '').toLowerCase().trim();
    const levelValue = levelInput && levelInput.value;
    const levelFilter = parseInt(levelValue, 10);
    const hasLevelFilter = !isNaN(levelFilter);

    const nodes = document.querySelectorAll('.node');

    nodes.forEach(function(el) {
      const name = (el.dataset.name || '').toLowerCase();
      const level = parseInt(el.dataset.level || '0', 10);

      let match = true;
      if (nameFilter && name.indexOf(nameFilter) === -1) {
        match = false;
      }
      if (hasLevelFilter && level > levelFilter) {
        match = false;
      }

      el.dataset.match = match ? '1' : '0';
    });

    nodes.forEach(function(el) {
      const type = el.dataset.type || 'file';
      let visible = (el.dataset.match === '1');

      if (type === 'dir' && !visible) {
        const details = el.closest('details');
        if (details) {
          const childMatch = details.querySelector('.node[data-match="1"]');
          if (childMatch) {
            visible = true;
          }
        }
      }

      let container = el.closest('li');
      if (!container) {
        if (type === 'dir') {
          container = el.closest('details') || el;
        } else {
          container = el;
        }
      }

      container.style.display = visible ? '' : 'none';
    });
  }

  function resetFilters() {
    if (nameInput) nameInput.value = '';
    if (levelInput) levelInput.value = '';

    const containers = document.querySelectorAll('details, li, .node');
    containers.forEach(function(el) {
      el.style.display = '';
      if (el.classList && el.classList.contains('node')) {
        el.dataset.match = '1';
      }
    });
  }

  if (applyBtn) applyBtn.addEventListener('click', applyFilters);
  if (resetBtn) resetBtn.addEventListener('click', resetFilters);
  if (nameInput) nameInput.addEventListener('input', applyFilters);
})();
</script>
"""
        )

        html_parts.append("</body></html>")

        with filepath.open("w", encoding="utf-8") as f:
            f.write("".join(html_parts))


def main():
    root = tk.Tk()
    app = WinDirScopeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
