import os
import threading
import csv
import datetime
import json
from pathlib import Path
from dataclasses import dataclass, field

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


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
                        child_path = Path(entry.path)
                        try:
                            child_node = _scan(child_path, level + 1)
                            node.children.append(child_node)
                            node.size += child_node.size
                        except (PermissionError, FileNotFoundError, OSError):
                            # On ignore ce qu'on ne peut pas lire
                            continue
                        finally:
                            if progress_callback:
                                progress_callback()
            except (PermissionError, FileNotFoundError, OSError):
                # Accès totalement refusé à ce dossier
                node.access_denied = True
            return node
        else:
            try:
                size = path.stat().st_size
            except (PermissionError, FileNotFoundError, OSError):
                size = 0
            node = Node(path=path, name=path.name, is_dir=False, size=size, level=level)

            ext = path.suffix.lower()
            if not ext:
                ext = "<sans extension>"
            ext_stats[ext] = ext_stats.get(ext, 0) + size

            if progress_callback:
                progress_callback()

            return node

    root_node = _scan(root_path, 0)
    return root_node, ext_stats


class WinDirScopeApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title("WinDirScope - Analyse d'espace disque")
        self.master.geometry("1000x650")

        self.root_node = None
        self.ext_stats = {}
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

        self._build_ui()

    # ---------- UI ----------

    def _build_ui(self):
        self._build_menu()

        # Barre supérieure : bouton + label + profondeur + barre de progression
        top_frame = ttk.Frame(self.master)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.btn_select = ttk.Button(
            top_frame,
            text="Analyser un dossier…",
            command=self.on_select_folder
        )
        self.btn_select.pack(side=tk.LEFT)

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

        # Split vertical : arbre en haut, extensions en bas
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

        # --- Tableau des extensions ---
        ext_frame = ttk.LabelFrame(paned, text="Répartition par extension")
        paned.add(ext_frame, weight=1)

        ext_columns = ("ext", "size", "percent")
        self.ext_tree = ttk.Treeview(
            ext_frame,
            columns=ext_columns,
            show="headings",
            height=6
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
            "WinDirScope\n"
            "Analyse et export d'utilisation disque (inspiré de WinDirStat, réécrit en Python/Tkinter).\n"
            "Les dossiers en rouge avec [ACCÈS REFUSÉ] sont ceux pour lesquels les droits sont insuffisants.\n"
            "L'export HTML permet de replier/déplier les dossiers dans un navigateur."
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
                self._populate_views()

    def _update_progress_ui(self):
        if self.progress_total <= 0:
            self.progress_var.set(0.0)
            return
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
        # On ne touche pas aux stats extensions, seulement à l'arbre
        self._clear_tree_view()
        self._populate_tree_view()

    # ---------- Gestion de l'affichage ----------

    def _clear_views(self):
        self._clear_tree_view()
        self.ext_tree.delete(*self.ext_tree.get_children())
        self.id_to_node.clear()
        self.id_counter = 0

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

            # Texte + indicateur visuel accès refusé
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

        # Nom de base par défaut : date + nom du dossier analysé
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        root_name = self.root_node.name or "analyse"
        default_name = f"WinDirScope_{timestamp}_{root_name}.csv"

        base_path_str = filedialog.asksaveasfilename(
            title="Exporter les résultats (nom de base du fichier)",
            defaultextension=".csv",
            filetypes=[
                ("Fichier CSV", "*.csv"),
                ("Fichier JSON", "*.json"),
                ("Fichier texte", "*.txt"),
                ("Page HTML", "*.html"),
                ("Page HTML (htm)", "*.htm"),
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
                self._export_tree_json(tree_file)
                self._export_ext_json(ext_file)
                exported_files = [tree_file, ext_file]
            elif suffix == ".txt":
                tree_file = base_path.with_name(base_path.stem + "_arborescence.txt")
                ext_file = base_path.with_name(base_path.stem + "_extensions.txt")
                self._export_tree_txt(tree_file)
                self._export_ext_txt(ext_file)
                exported_files = [tree_file, ext_file]
            elif suffix in (".html", ".htm"):
                html_file = base_path
                self._export_html(html_file)
                exported_files = [html_file]
            else:  # CSV par défaut
                tree_file = base_path.with_name(base_path.stem + "_arborescence.csv")
                ext_file = base_path.with_name(base_path.stem + "_extensions.csv")
                self._export_tree_csv(tree_file)
                self._export_ext_csv(ext_file)
                exported_files = [tree_file, ext_file]
        except Exception as e:
            messagebox.showerror("Erreur d'export", f"Impossible d'exporter les résultats : {e}")
            return

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
            for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
                percent = (size / total_ext_size) * 100
                writer.writerow([
                    ext,
                    size,
                    human_size(size),
                    f"{percent:.4f}",
                ])

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
            items.append({
                "extension": ext,
                "size_bytes": size,
                "size_human": human_size(size),
                "percent_total": percent,
            })
        with filepath.open("w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)

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

    # --- HTML ---

    def _export_html(self, filepath: Path):
        """Export en page HTML avec dossier repliables/dépliables."""
        def esc(s: str) -> str:
            return (
                s.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;")
            )

        total_size = self.root_node.size or 1
        total_ext_size = sum(self.ext_stats.values()) or 1

        def node_to_html(node: Node) -> str:
            percent = (node.size / total_size) * 100
            name = esc(node.name)
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

            # Ligne principale (texte)
            line = (
                f'<span class="name">{name}</span> '
                f'<span class="meta">({esc(info)})</span><br>'
                f'<span class="path">{path}</span>'
            )

            if node.is_dir:
                # Dossiers : <details><summary>...</summary><ul>...</ul></details>
                open_attr = " open" if node.level <= 1 else ""
                classes = "dir denied" if denied else "dir"
                html = [f'<details{open_attr}><summary class="{classes}">{line}</summary>']
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
                classes = "file denied" if denied else "file"
                return f'<div class="{classes}">{line}</div>'

        html_parts = []
        html_parts.append("<!DOCTYPE html>")
        html_parts.append("<html lang='fr'>")
        html_parts.append("<head>")
        html_parts.append("<meta charset='utf-8'>")
        html_parts.append(f"<title>WinDirScope - Rapport {esc(self.root_node.name)}</title>")
        html_parts.append("<style>")
        html_parts.append("""
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
                display: flex;
                gap: 16px;
                padding: 16px;
            }
            section {
                background: #ffffff;
                border-radius: 6px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.08);
                padding: 12px 16px;
                box-sizing: border-box;
            }
            section.tree {
                flex: 2;
                max-height: 80vh;
                overflow: auto;
            }
            section.ext {
                flex: 1;
                max-height: 80vh;
                overflow: auto;
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
        """)
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
        # Arbre
        html_parts.append("<section class='tree'>")
        html_parts.append("<h2>Arborescence</h2>")
        html_parts.append(node_to_html(self.root_node))
        html_parts.append("</section>")

        # Extensions
        html_parts.append("<section class='ext'>")
        html_parts.append("<h2>Répartition par extension</h2>")
        html_parts.append("<table>")
        html_parts.append("<thead><tr><th>Extension</th><th>Taille lisible</th><th>Taille (octets)</th><th>% du total</th></tr></thead>")
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
        html_parts.append("</main>")

        html_parts.append("<footer>")
        html_parts.append(f"Rapport généré par WinDirScope le {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}.")
        html_parts.append("</footer>")

        html_parts.append("</body></html>")

        with filepath.open("w", encoding="utf-8") as f:
            f.write("".join(html_parts))


def main():
    root = tk.Tk()
    app = WinDirScopeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
