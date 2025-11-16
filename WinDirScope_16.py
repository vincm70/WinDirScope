import os
import sys
import subprocess
import threading
import csv
import datetime
import json
import shutil
from pathlib import Path
from dataclasses import dataclass, field

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

APP_VERSION = "1.2.0"  # version avec menu clic droit + progression corrigée


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
            # Pour les fichiers, on ne touche plus à la progression : déjà comptés comme entrée du dossier parent
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

        # Profondeur max d'affichage / analyse visible
        self.max_level_var = tk.IntVar(value=5)

        # Pour le menu contextuel
        self._context_item_id = None
        self._context_node = None

        self._build_ui()

    # ---------- UI ----------

    def _build_ui(self):
        self._build_menu()

        # Barre supérieure : boutons + label + profondeur + barre de progression
        top_frame = ttk.Frame(self.master)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.btn_select = ttk.Button(
            top_frame,
            text="Analyser un dossier…",
            command=self.on_select_folder
        )
        self.btn_select.pack(side=tk.LEFT)

        self.btn_export = ttk.Button(
            top_frame,
            text="Exporter…",
            command=self.on_export_results,
            state="disabled"  # activé après une analyse réussie
        )
        self.btn_export.pack(side=tk.LEFT, padx=(5, 0))

        self.lbl_status = ttk.Label(top_frame, text="Aucun dossier analysé.")
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # Choix profondeur max
        self.lbl_level = ttk.Label(top_frame, text="Profondeur analysée (niveaux) :")
        self.lbl_level.pack(side=tk.LEFT, padx=(10, 2))

        self.spin_level = ttk.Spinbox(
            top_frame,
            from_=1,
            to=20,
            textvariable=self.max_level_var,
            width=3,
            command=self.on_change_max_level
        )
        self.spin_level.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(
            top_frame,
            orient="horizontal",
            mode="determinate",
            variable=self.progress_var,
            maximum=100
        )
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        # Split vertical : 3 zones (arbre, extensions, top 100)
        paned = ttk.PanedWindow(self.master, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Arbre principal ---
        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3)

        columns = ("level", "size", "percent")
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="tree headings"
        )
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

        # Style pour les dossiers en accès refusé
        self.tree.tag_configure("denied", foreground="red")

        # Menu contextuel sur l’arborescence
        self.tree_menu = tk.Menu(self.master, tearoff=False)
        self.tree_menu.add_command(label="Ouvrir la cible", command=self.cmd_open_target)
        self.tree_menu.add_command(label="Ouvrir le dossier contenant", command=self.cmd_open_containing)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Renommer…", command=self.cmd_rename_node)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="Supprimer définitivement…", command=self.cmd_delete_node)

        # Clic droit
        self.tree.bind("<Button-3>", self.on_tree_right_click)

        # --- Tableau des extensions ---
        ext_frame = ttk.LabelFrame(paned, text="Répartition par extension")
        paned.add(ext_frame, weight=2)

        ext_columns = ("ext", "size", "percent")
        self.ext_tree = ttk.Treeview(
            ext_frame,
            columns=ext_columns,
            show="headings",
            height=4
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
            top_frame,
            columns=top_columns,
            show="headings",
            height=6
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

    def _build_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Analyser un dossier…", command=self.on_select_folder)
        file_menu.add_command(label="Exporter les résultats…", command=self.on_export_results)
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
                "- Menu clic droit (ouvrir, renommer, supprimer)\n"
                "- Export CSV / JSON / TXT / HTML (rapport HTML interactif)\n"
                "- Indication visuelle des dossiers en accès refusé.\n"
            )
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
            messagebox.showerror("Erreur", "Ce dossier n'existe pas ou n'est pas accessible.")
            return

        # Pré-comptage des entrées pour la progression
        self.current_scan_path = path
        self.progress_total = count_entries(path)
        self.progress_current = 0
        self.progress_var.set(0.0)

        # Lancer le scan dans un thread
        self.scan_running = True
        self.btn_export.config(state="disabled")
        self.lbl_status.config(text=f"Analyse en cours : {folder}")
        self._clear_views()

        self.scan_thread = threading.Thread(
            target=self._scan_worker,
            args=(path,),
            daemon=True
        )
        self.scan_thread.start()
        self.master.after(200, self._poll_scan_thread)

    def _scan_worker(self, path: Path):
        """Fonction exécutée dans un thread séparé."""
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
        """Appelée dans le thread de scan pour incrémenter le compteur."""
        self.progress_current += 1

    def _poll_scan_thread(self):
        if self.scan_thread is None:
            return

        # Met à jour la barre de progression
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
                # Calcul du Top 100 fichiers
                self._compute_top_files()
                self._populate_views()
                self.btn_export.config(state="normal")

    def _update_progress_ui(self):
        if self.progress_total <= 0:
            self.progress_var.set(0.0)
            return
        # On borne à 100 % au cas où, mais la logique de comptage est maintenant alignée
        percent = min(100.0, (self.progress_current / self.progress_total) * 100.0)
        self.progress_var.set(percent)
        if self.current_scan_path:
            self.lbl_status.config(
                text=f"Analyse en cours : {self.current_scan_path} "
                     f"({self.progress_current}/{self.progress_total} éléments, {percent:5.1f} %)"
            )

    def on_change_max_level(self):
        """Re-génère la vue arbre selon la profondeur choisie."""
        if not self.root_node or self.scan_running:
            return
        self._clear_tree_view()
        self._populate_tree_view()

    # ---------- Gestion de l'affichage ----------

    def _clear_views(self):
        self._clear_tree_view()
        self.ext_tree.delete(*self.ext_tree.get_children())
        self.top_tree.delete(*self.top_tree.get_children())
        self.id_to_node.clear()
        self.id_counter = 0
        self.top_files = []

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

        total_size = self.root_node.size or 1  # éviter division par zéro

        try:
            max_level = int(self.max_level_var.get())
        except (TypeError, ValueError):
            max_level = 5

        def add_node_to_tree(parent_id, node: Node):
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
                tags=tags
            )
            if node.is_dir:
                for child in sorted(node.children, key=lambda n: n.size, reverse=True):
                    add_node_to_tree(tree_id, child)

        add_node_to_tree("", self.root_node)
        # Ouvrir le premier niveau
        first_child = self.tree.get_children()
        if first_child:
            self.tree.item(first_child[0], open=True)

    def _populate_ext_view(self):
        self.ext_tree.delete(*self.ext_tree.get_children())
        total_ext_size = sum(self.ext_stats.values()) or 1
        for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
            percent = (size / total_ext_size) * 100
            self.ext_tree.insert(
                "",
                "end",
                values=(ext, human_size(size), f"{percent:5.2f} %")
            )

    def _populate_top_files_view(self):
        self.top_tree.delete(*self.top_tree.get_children())
        for row in self.top_files:
            self.top_tree.insert(
                "",
                "end",
                values=(
                    row["size_human"],
                    f"{row['percent_total']:.2f} %",
                    row["path"],
                )
            )

    # ---------- Calcul Top 100 fichiers ----------

    def _compute_top_files(self):
        """Calcule la liste des 100 fichiers les plus volumineux."""
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
                files.append({
                    "path": str(node.path),
                    "name": node.name,
                    "size_bytes": node.size,
                    "size_human": human_size(node.size),
                    "percent_total": percent,
                    "level": node.level,
                })

        visit(self.root_node)
        files.sort(key=lambda r: r["size_bytes"], reverse=True)
        self.top_files = files[:100]

    # ---------- Menu contextuel arbre ----------

    def on_tree_right_click(self, event):
        """Affiche le menu contextuel sur l’élément sous le clic droit."""
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
        """Ouvre la cible (fichier ou dossier) avec l’appli par défaut."""
        item_id, node = self._get_context_node()
        if not node:
            return
        path = node.path
        if not path.exists():
            messagebox.showerror("Chemin introuvable", f"Le chemin n'existe plus :\n{path}")
            return
        open_file_in_default_app(path)

    def cmd_open_containing(self):
        """Ouvre le dossier contenant l’élément."""
        item_id, node = self._get_context_node()
        if not node:
            return
        path = node.path
        if node.is_dir:
            folder = path
        else:
            folder = path.parent
        if not folder.exists():
            messagebox.showerror("Chemin introuvable", f"Le dossier n'existe plus :\n{folder}")
            return
        open_file_in_default_app(folder)

    def cmd_rename_node(self):
        """Renomme un fichier ou un dossier (et met à jour les chemins enfants)."""
        item_id, node = self._get_context_node()
        if not node:
            return

        old_name = node.name
        new_name = simpledialog.askstring(
            "Renommer",
            f"Nouveau nom pour :\n{node.path}",
            initialvalue=old_name,
            parent=self.master
        )
        if not new_name or new_name == old_name:
            return

        old_path = node.path
        new_path = old_path.with_name(new_name)

        if new_path.exists():
            messagebox.showerror("Existe déjà", f"Un élément portant ce nom existe déjà :\n{new_path}")
            return

        try:
            os.rename(old_path, new_path)
        except Exception as e:
            messagebox.showerror("Erreur de renommage", f"Impossible de renommer :\n{e}")
            return

        # Mise à jour du node
        node.path = new_path
        node.name = new_name

        # Si dossier, il faut mettre à jour les chemins de tous les enfants
        if node.is_dir:
            self._update_child_paths(node, old_path, new_path)

        # Mise à jour de l'affichage
        self.tree.item(item_id, text=new_name)

        messagebox.showinfo(
            "Renommage effectué",
            "Le renommage a été effectué sur le système de fichiers.\n"
            "Les statistiques affichées ne sont pas recalculées automatiquement.\n"
            "Relancez une analyse si nécessaire."
        )

    def _update_child_paths(self, parent_node: Node, old_root: Path, new_root: Path):
        """Met à jour les paths des enfants après renommage d’un dossier."""
        for child in parent_node.children:
            try:
                rel = child.path.relative_to(old_root)
                child.path = new_root / rel
            except ValueError:
                # Si pour une raison quelconque le chemin n'est pas sous old_root
                pass
            if child.is_dir:
                self._update_child_paths(child, old_root, new_root)

    def cmd_delete_node(self):
        """Supprime définitivement un fichier ou un dossier."""
        item_id, node = self._get_context_node()
        if not node:
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
            icon=messagebox.WARNING
        ):
            return

        try:
            if node.is_dir:
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception as e:
            messagebox.showerror("Erreur de suppression", f"Impossible de supprimer :\n{e}")
            return

        # Retirer du modèle et de l'arbre
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
            "Relancez une analyse si nécessaire."
        )

    # ---------- Flatten commun ----------

    def _flatten_tree(self):
        """Retourne une liste de lignes à exporter pour l'arborescence."""
        rows = []
        total_size = self.root_node.size or 1

        def visit(node: Node):
            percent = (node.size / total_size) * 100
            rows.append({
                "path": str(node.path),
                "name": node.name,
                "level": node.level,
                "type": "dossier" if node.is_dir else "fichier",
                "size_bytes": node.size,
                "size_human": human_size(node.size),
                "percent_total": percent,
                "access_denied": node.access_denied,
            })
            for child in node.children:
                visit(child)

        visit(self.root_node)
        return rows

    # ---------- Export ----------

    def on_export_results(self):
        if not self.root_node:
            messagebox.showwarning(
                "Aucun résultat",
                "Aucun dossier n'a encore été analysé. Lancez une analyse avant d'exporter."
            )
            return

        # Assurer que le Top 100 est à jour
        self._compute_top_files()

        # Nom par défaut : HTML, avec nettoyage des caractères interdits
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
            initialfile=default_name
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
            else:  # CSV par défaut
                tree_file = base_path.with_name(base_path.stem + "_arborescence.csv")
                ext_file = base_path.with_name(base_path.stem + "_extensions.csv")
                top_file = base_path.with_name(base_path.stem + "_top100.csv")
                self._export_tree_csv(tree_file)
                self._export_ext_csv(ext_file)
                self._export_top_csv(top_file)
                exported_files = [tree_file, ext_file, top_file]
        except Exception as e:
            messagebox.showerror("Erreur d'export", f"Impossible d'exporter les résultats : {e}")
            return

        # Ouvrir automatiquement le rapport principal (premier fichier exporté)
        if exported_files:
            open_file_in_default_app(exported_files[0])

        msg = "Résultats exportés :\n" + "\n".join(f"- {p}" for p in exported_files)
        messagebox.showinfo("Export terminé", msg)

    # --- CSV ---

    def _export_tree_csv(self, filepath: Path):
        rows = self._flatten_tree()
        with filepath.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow([
                "Chemin complet",
                "Nom",
                "Niveau",
                "Type",
                "Taille (octets)",
                "Taille lisible",
                "% du total",
                "Accès refusé"
            ])
            for row in rows:
                writer.writerow([
                    row["path"],
                    row["name"],
                    row["level"],
                    row["type"],
                    row["size_bytes"],
                    row["size_human"],
                    f"{row['percent_total']:.4f}",
                    "Oui" if row["access_denied"] else "Non",
                ])

    def _export_ext_csv(self, filepath: Path):
        total_ext_size = sum(self.ext_stats.values()) or 1
        with filepath.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["Extension", "Taille totale (octets)", "Taille lisible", "% du total"])
        ...
