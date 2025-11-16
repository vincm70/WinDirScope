import os
import threading
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


def human_size(num_bytes: int) -> str:
    """Convertit un nombre d'octets en format lisible."""
    for unit in ["o", "Ko", "Mo", "Go", "To"]:
        if num_bytes < 1024 or unit == "To":
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} To"


def scan_directory(root_path: Path):
    """
    Scan récursif du dossier.
    Retourne (root_node, stats_extensions)
    stats_extensions = {".txt": taille_totale, ...}
    """
    ext_stats = {}

    def _scan(path: Path) -> Node:
        if path.is_dir():
            node = Node(path=path, name=path.name or str(path), is_dir=True, size=0)
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        child_path = Path(entry.path)
                        try:
                            child_node = _scan(child_path)
                            node.children.append(child_node)
                            node.size += child_node.size
                        except (PermissionError, FileNotFoundError, OSError):
                            # On ignore ce qu'on ne peut pas lire
                            continue
            except (PermissionError, FileNotFoundError, OSError):
                pass
            return node
        else:
            try:
                size = path.stat().st_size
            except (PermissionError, FileNotFoundError, OSError):
                size = 0
            node = Node(path=path, name=path.name, is_dir=False, size=size)

            ext = path.suffix.lower()
            if not ext:
                ext = "<sans extension>"
            ext_stats[ext] = ext_stats.get(ext, 0) + size

            return node

    root_node = _scan(root_path)
    return root_node, ext_stats


class PyDirStatApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title("PyDirStat - Analyse d'espace disque (simplifié)")
        self.master.geometry("900x600")

        self.root_node = None
        self.ext_stats = {}
        self.scan_thread = None
        self.scan_running = False

        self.id_counter = 0
        self.id_to_node = {}

        self._build_ui()

    # ---------- UI ----------

    def _build_ui(self):
        self._build_menu()

        # Barre supérieure : bouton + label
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

        # Split vertical : arbre en haut, extensions en bas
        paned = ttk.PanedWindow(self.master, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Arbre principal ---
        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3)

        columns = ("size", "percent")
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="tree headings"
        )
        self.tree.heading("#0", text="Dossier / Fichier")
        self.tree.heading("size", text="Taille")
        self.tree.heading("percent", text="Pourcentage")

        self.tree.column("#0", width=400, anchor=tk.W)
        self.tree.column("size", width=120, anchor=tk.E)
        self.tree.column("percent", width=100, anchor=tk.E)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

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
            "PyDirStat (version simplifiée)\n"
            "Réécrit en Python/Tkinter à partir de l'idée de WinDirStat."
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
            messagebox.showerror("Erreur", "Ce dossier n'existe pas.")
            return

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
            root_node, ext_stats = scan_directory(path)
            self.root_node = root_node
            self.ext_stats = ext_stats
        except Exception as e:
            # On stocke l'erreur pour l'afficher ensuite dans le thread principal
            self.root_node = None
            self.ext_stats = {}
            self._scan_error = e
        else:
            self._scan_error = None

    def _poll_scan_thread(self):
        if self.scan_thread is None:
            return

        if self.scan_thread.is_alive():
            # On peut mettre une petite animation éventuellement
            self.master.after(200, self._poll_scan_thread)
        else:
            self.scan_running = False
            if self._scan_error:
                messagebox.showerror("Erreur d'analyse", str(self._scan_error))
                self.lbl_status.config(text="Erreur lors de l'analyse.")
            else:
                self.lbl_status.config(text=f"Analyse terminée : {self.root_node.path}")
                self._populate_views()

    # ---------- Gestion de l'affichage ----------

    def _clear_views(self):
        self.tree.delete(*self.tree.get_children())
        self.ext_tree.delete(*self.ext_tree.get_children())
        self.id_to_node.clear()
        self.id_counter = 0

    def _next_id(self):
        self.id_counter += 1
        return f"node_{self.id_counter}"

    def _populate_views(self):
        if not self.root_node:
            return

        total_size = self.root_node.size or 1  # éviter division par zéro
        # Arbre
        def add_node_to_tree(parent_id, node: Node):
            tree_id = self._next_id()
            self.id_to_node[tree_id] = node
            percent = (node.size / total_size) * 100
            self.tree.insert(
                parent_id,
                "end",
                iid=tree_id,
                text=node.name,
                values=(human_size(node.size), f"{percent:5.2f} %")
            )
            if node.is_dir:
                for child in sorted(node.children, key=lambda n: n.size, reverse=True):
                    add_node_to_tree(tree_id, child)

        add_node_to_tree("", self.root_node)
        self.tree.item(self.tree.get_children()[0], open=True)

        # Extensions
        total_ext_size = sum(self.ext_stats.values()) or 1
        for ext, size in sorted(self.ext_stats.items(), key=lambda kv: kv[1], reverse=True):
            percent = (size / total_ext_size) * 100
            self.ext_tree.insert(
                "",
                "end",
                values=(ext, human_size(size), f"{percent:5.2f} %")
            )


def main():
    root = tk.Tk()
    app = PyDirStatApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
