"""
@author João Gbriel de Almeida
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from src.pst_reader import PstReader
from src.models import PstFolder, PstEmail

try:
    from tkhtmlview import HTMLLabel  # type: ignore
except Exception:  # pragma: no cover
    HTMLLabel = None  # type: ignore


class AppUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.reader: Optional[PstReader] = None

        self._build_menu()
        self._build_layout()

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Abrir PST...", command=self._on_open_pst)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit)
        menu_bar.add_cascade(label="Arquivo", menu=file_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Sobre", command=self._on_about)
        menu_bar.add_cascade(label="Ajuda", menu=help_menu)

        self.root.config(menu=menu_bar)

    def _build_layout(self) -> None:
        self.paned = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # Esquerda: árvore de pastas
        left_frame = ttk.Frame(self.paned)
        self.tree = ttk.Treeview(left_frame, columns=("name",), show="tree")
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_folder_selected)
        self.paned.add(left_frame, weight=1)

        # Centro: lista de mensagens
        center_frame = ttk.Frame(self.paned)
        columns = ("assunto", "remetente", "data")
        self.msg_list = ttk.Treeview(center_frame, columns=columns, show="headings")
        self.msg_list.heading("assunto", text="Assunto")
        self.msg_list.heading("remetente", text="Remetente")
        self.msg_list.heading("data", text="Data")
        self.msg_list.pack(fill=tk.BOTH, expand=True)
        self.msg_list.bind("<<TreeviewSelect>>", self._on_message_selected)
        self.paned.add(center_frame, weight=2)

        # Direita: preview + anexos
        right_frame = ttk.Frame(self.paned)
        # Cabeçalhos breves
        self.header_text = tk.Text(right_frame, height=6, wrap=tk.WORD)
        self.header_text.pack(fill=tk.X)

        # Corpo: HTML se disponível
        self.preview_container = ttk.Frame(right_frame)
        self.preview_container.pack(fill=tk.BOTH, expand=True)
        if HTMLLabel is not None:
            self.html_preview = HTMLLabel(self.preview_container, html="")
            self.html_preview.pack(fill=tk.BOTH, expand=True)
            self.text_preview = None
        else:
            self.html_preview = None
            self.text_preview = tk.Text(self.preview_container, wrap=tk.WORD)
            self.text_preview.pack(fill=tk.BOTH, expand=True)

        # Anexos + ações
        attachments_bar = ttk.Frame(right_frame)
        attachments_bar.pack(fill=tk.X)
        ttk.Label(attachments_bar, text="Anexos:").pack(side=tk.LEFT)
        self.attach_list = tk.Listbox(attachments_bar, height=4)
        self.attach_list.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(attachments_bar, text="Salvar Anexos", command=self._save_attachments).pack(side=tk.RIGHT)

        self.paned.add(right_frame, weight=3)

        # Barra de busca
        bottom = ttk.Frame(self.root)
        bottom.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(bottom, text="Buscar:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(bottom, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(bottom, text="Aplicar", command=self._apply_search).pack(side=tk.LEFT)

    def _on_about(self) -> None:
        messagebox.showinfo("Sobre", "Leitor de PST em Python (Tkinter)")

    def _on_open_pst(self) -> None:
        path = filedialog.askopenfilename(title="Escolher arquivo PST", filetypes=[("Outlook PST", "*.pst"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.reader = PstReader()
            self.reader.open(path)
            self._load_tree()
            self._clear_messages()
            self._clear_preview()
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Erro ao abrir PST", str(exc))

    def _load_tree(self) -> None:
        if not self.reader:
            return
        self.tree.delete(*self.tree.get_children())
        for folder in self.reader.get_root_folders():
            node = self.tree.insert("", tk.END, iid=folder.id, text=folder.name)
            self._insert_children(node, folder)

    def _insert_children(self, node_id: str, folder: PstFolder) -> None:
        for child in folder.children:
            child_id = self.tree.insert(node_id, tk.END, iid=child.id, text=child.name)
            self._insert_children(child_id, child)

    def _on_folder_selected(self, _event=None) -> None:
        if not self.reader:
            return
        selected = self.tree.selection()
        if not selected:
            return
        folder_id = selected[0]
        self._populate_messages(folder_id)

    def _populate_messages(self, folder_id: str) -> None:
        self._clear_messages()
        if not self.reader:
            return
        for msg in self.reader.list_messages(folder_id):
            self.msg_list.insert("", tk.END, iid=msg.id, values=(msg.subject, msg.sender, msg.date or ""))

    def _on_message_selected(self, _event=None) -> None:
        if not self.reader:
            return
        selected = self.msg_list.selection()
        if not selected:
            return
        msg_id = selected[0]
        msg = self.reader.get_message(msg_id)
        self._show_message(msg)
        # carregar anexos
        self._load_attachments(msg_id)

    def _show_message(self, msg: PstEmail) -> None:
        self._clear_preview()
        headers = [
            f"Assunto: {msg.subject}",
            f"De: {msg.sender}",
            f"Para: {msg.to}",
            f"Cc: {msg.cc}",
            f"Data: {msg.date or ''}",
        ]
        self.header_text.insert("1.0", "\n".join(headers))

        if self.html_preview is not None and msg.body_html:
            self.html_preview.set_html(msg.body_html)
        else:
            text = msg.body_text or msg.body_html or "(sem corpo)"
            if self.text_preview is not None:
                self.text_preview.insert("1.0", text)

    def _load_attachments(self, msg_id: str) -> None:
        self.attach_list.delete(0, tk.END)
        if not self.reader:
            return
        try:
            names = self.reader.get_attachments(msg_id)
        except Exception:
            names = []
        for name in names:
            self.attach_list.insert(tk.END, name)

    def _save_attachments(self) -> None:
        if not self.reader:
            return
        selected = self.msg_list.selection()
        if not selected:
            return
        msg_id = selected[0]
        out_dir = filedialog.askdirectory(title="Selecionar pasta para salvar anexos")
        if not out_dir:
            return
        try:
            saved = self.reader.save_attachments(msg_id, out_dir)
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Salvar Anexos", str(exc))
            return
        messagebox.showinfo("Salvar Anexos", f"{len(saved)} anexo(s) salvo(s).")

    def _apply_search(self) -> None:
        term = (self.search_var.get() or "").strip().lower()
        selected = self.tree.selection()
        if not selected or not self.reader:
            return
        folder_id = selected[0]
        self._clear_messages()
        for msg in self.reader.list_messages(folder_id):
            blob = " ".join([msg.subject or "", msg.sender or "", (msg.body_text or ""), (msg.body_html or "")]).lower()
            if term in blob:
                self.msg_list.insert("", tk.END, iid=msg.id, values=(msg.subject, msg.sender, msg.date or ""))

    def _clear_messages(self) -> None:
        self.msg_list.delete(*self.msg_list.get_children())

    def _clear_preview(self) -> None:
        self.header_text.delete("1.0", tk.END)
        if self.html_preview is not None:
            self.html_preview.set_html("")
        if self.text_preview is not None:
            self.text_preview.delete("1.0", tk.END)
        self.attach_list.delete(0, tk.END)
