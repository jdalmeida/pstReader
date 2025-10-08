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

        # Tema ttk
        try:
            style = ttk.Style()
            if "vista" in style.theme_names():
                style.theme_use("vista")
        except Exception:
            pass

        self._build_menu()
        self._build_toolbar()
        self._build_layout()
        self._build_statusbar()

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Abrir PST...", command=self._on_open_pst)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.root.quit)
        menu_bar.add_cascade(label="Arquivo", menu=file_menu)

        action_menu = tk.Menu(menu_bar, tearoff=0)
        action_menu.add_command(label="Exportar EML", command=self._export_selected_eml)
        action_menu.add_command(label="Salvar Anexos", command=self._save_attachments)
        menu_bar.add_cascade(label="Ações", menu=action_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Sobre", command=self._on_about)
        menu_bar.add_cascade(label="Ajuda", menu=help_menu)

        self.root.config(menu=menu_bar)

    def _build_toolbar(self) -> None:
        tb = ttk.Frame(self.root)
        tb.pack(fill=tk.X, side=tk.TOP)
        ttk.Button(tb, text="Abrir PST", command=self._on_open_pst).pack(side=tk.LEFT, padx=4, pady=2)
        ttk.Button(tb, text="Exportar EML", command=self._export_selected_eml).pack(side=tk.LEFT, padx=4, pady=2)
        ttk.Button(tb, text="Salvar Anexos", command=self._save_attachments).pack(side=tk.LEFT, padx=4, pady=2)

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
        self.msg_list.column("assunto", width=400, anchor=tk.W)
        self.msg_list.column("remetente", width=200, anchor=tk.W)
        self.msg_list.column("data", width=150, anchor=tk.W)
        self.msg_list.pack(fill=tk.BOTH, expand=True)
        self.msg_list.bind("<<TreeviewSelect>>", self._on_message_selected)
        self._build_msg_context_menu()
        self.paned.add(center_frame, weight=2)

        # Direita: preview + anexos
        right_frame = ttk.Frame(self.paned)
        self.header_text = tk.Text(right_frame, height=6, wrap=tk.WORD)
        self.header_text.pack(fill=tk.X)

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

    def _build_statusbar(self) -> None:
        sb = ttk.Frame(self.root)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="Pronto")
        ttk.Label(sb, textvariable=self.status_var, anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _build_msg_context_menu(self) -> None:
        self.msg_menu = tk.Menu(self.root, tearoff=0)
        self.msg_menu.add_command(label="Exportar EML", command=self._export_selected_eml)
        self.msg_menu.add_command(label="Salvar Anexos", command=self._save_attachments)
        self.msg_list.bind("<Button-3>", self._on_msg_right_click)

    def _on_msg_right_click(self, event) -> None:
        try:
            row = self.msg_list.identify_row(event.y)
            if row:
                self.msg_list.selection_set(row)
            self.msg_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.msg_menu.grab_release()

    def _set_busy(self, busy: bool) -> None:
        self.root.config(cursor="watch" if busy else "")
        self.root.update_idletasks()
        if not busy:
            self.status_var.set("Pronto")

    def _on_about(self) -> None:
        messagebox.showinfo(
            "Sobre",
            "Leitor de PST em Python (Tkinter)\n\n"
            "Autor: João Gbriel de Almeida\n"
            "Licença: MIT\n"
            "\nEste software é fornecido \"no estado em que se encontra\", sem garantias.",
        )

    def _on_open_pst(self) -> None:
        path = filedialog.askopenfilename(title="Escolher arquivo PST", filetypes=[("Outlook PST", "*.pst"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self._set_busy(True)
            self.status_var.set("Abrindo PST...")
            self.reader = PstReader()
            self.reader.open(path)
            self._load_tree()
            self._clear_messages()
            self._clear_preview()
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Erro ao abrir PST", str(exc))
        finally:
            self._set_busy(False)

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
        self.status_var.set(f"{len(self.msg_list.get_children())} mensagem(ns)")

    def _on_message_selected(self, _event=None) -> None:
        if not self.reader:
            return
        selected = self.msg_list.selection()
        if not selected:
            return
        msg_id = selected[0]
        self._set_busy(True)
        try:
            msg = self.reader.get_message(msg_id)
            self._show_message(msg)
            self._load_attachments(msg_id)
        finally:
            self._set_busy(False)

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
            self._set_busy(True)
            self.status_var.set("Salvando anexos...")
            saved = self.reader.save_attachments(msg_id, out_dir)
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Salvar Anexos", str(exc))
            return
        finally:
            self._set_busy(False)
        messagebox.showinfo("Salvar Anexos", f"{len(saved)} anexo(s) salvo(s).")

    def _export_selected_eml(self) -> None:
        if not self.reader:
            return
        selected = self.msg_list.selection()
        if not selected:
            return
        msg_id = selected[0]
        path = filedialog.asksaveasfilename(title="Salvar como .eml", defaultextension=".eml", filetypes=[("EML", "*.eml"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self._set_busy(True)
            self.status_var.set("Exportando EML...")
            self.reader.export_eml(msg_id, path)
        except Exception as exc:  # pragma: no cover
            messagebox.showerror("Exportar EML", str(exc))
        finally:
            self._set_busy(False)

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
        self.status_var.set(f"{len(self.msg_list.get_children())} resultado(s)")

    def _clear_messages(self) -> None:
        self.msg_list.delete(*self.msg_list.get_children())

    def _clear_preview(self) -> None:
        self.header_text.delete("1.0", tk.END)
        if self.html_preview is not None:
            self.html_preview.set_html("")
        if self.text_preview is not None:
            self.text_preview.delete("1.0", tk.END)
        self.attach_list.delete(0, tk.END)
