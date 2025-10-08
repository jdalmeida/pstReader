"""
@author João Gbriel de Almeida
"""

from __future__ import annotations

from typing import Dict, List, Tuple
from pathlib import Path
import os
import mimetypes
import re

from src.models import PstEmail, PstFolder

try:
    import puremagic  # type: ignore
except Exception:  # pragma: no cover
    puremagic = None  # type: ignore


class PypffAdapter:
    def __init__(self) -> None:
        self._pff = None  # type: ignore
        self._file = None  # type: ignore
        self._folder_index: Dict[str, object] = {}

    def _normalize_path(self, path: str) -> str:
        try:
            p = Path(path).resolve(strict=False)
        except Exception:
            p = Path(path)
        norm = str(p).replace("/", "\\")
        if not norm.startswith("\\\\?\\") and len(norm) >= 240:
            norm = "\\\\?\\" + norm
        return norm

    def open(self, path: str) -> None:
        try:
            import pypff  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise RuntimeError("pypff não disponível") from exc
        self._pff = pypff
        self._file = pypff.file()
        norm_path = self._normalize_path(path)
        self._file.open(norm_path)
        self._index()

    def _index(self) -> None:
        root = self._file.get_root_folder()
        self._folder_index.clear()

        def walk(folder_obj) -> PstFolder:
            folder_id = str(id(folder_obj))
            name = getattr(folder_obj, "name", None) or getattr(folder_obj, "_name", "Pasta")
            children_py: List[PstFolder] = []
            try:
                count = folder_obj.number_of_sub_folders
            except Exception:
                count = getattr(folder_obj, "get_number_of_sub_folders", lambda: 0)()
            for i in range(count or 0):
                try:
                    child = folder_obj.get_sub_folder(i)
                except Exception:
                    continue
                child_model = walk(child)
                children_py.append(child_model)

            node = PstFolder(id=folder_id, name=name, children=children_py)
            self._folder_index[folder_id] = folder_obj
            return node

        self._root_nodes: List[PstFolder] = []
        try:
            sub_count = root.number_of_sub_folders
        except Exception:
            sub_count = getattr(root, "get_number_of_sub_folders", lambda: 0)()
        for i in range(sub_count or 0):
            try:
                f = root.get_sub_folder(i)
            except Exception:
                continue
            self._root_nodes.append(walk(f))

    # Public API
    def get_root_folders(self) -> List[PstFolder]:
        return list(self._root_nodes)

    def list_messages(self, folder_id: str) -> List[PstEmail]:
        folder_obj = self._folder_index.get(folder_id)
        if not folder_obj:
            return []
        emails: List[PstEmail] = []
        try:
            mcount = folder_obj.number_of_sub_messages
        except Exception:
            mcount = getattr(folder_obj, "get_number_of_sub_messages", lambda: 0)()
        for j in range(mcount or 0):
            try:
                msg = folder_obj.get_sub_message(j)
            except Exception:
                continue
            model = self._to_model_preview(msg)
            model.id = f"{folder_id}:{j}"
            emails.append(model)
        return emails

    def _resolve_message(self, composite_id: str):
        try:
            folder_id, idx_str = composite_id.rsplit(":", 1)
            idx = int(idx_str)
        except Exception as exc:
            raise KeyError("Mensagem não encontrada") from exc
        folder_obj = self._folder_index.get(folder_id)
        if not folder_obj:
            raise KeyError("Mensagem não encontrada")
        try:
            return folder_obj.get_sub_message(idx)
        except Exception as exc:
            raise KeyError("Mensagem não encontrada") from exc

    def get_message(self, msg_id: str) -> PstEmail:
        msg = self._resolve_message(msg_id)
        return self._to_model_full(msg, msg_id)

    def export_eml(self, msg_id: str, out_path: str) -> None:
        from src.utils.exporters import build_eml

        model = self.get_message(msg_id)
        content = build_eml(model)
        with open(out_path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(content)

    def _is_embedded_message(self, attachment) -> bool:
        for attr in ("is_embedded_message", "get_is_embedded_message"):
            try:
                v = getattr(attachment, attr)
                v = v() if callable(v) else v
                if isinstance(v, bool):
                    return v
            except Exception:
                continue
        try:
            _ = getattr(attachment, "get_embedded_message")
            return True
        except Exception:
            return False

    def _sanitize_filename(self, name: str) -> str:
        name = name.strip().replace("\n", " ").replace("\r", " ")
        name = re.sub(r"[<>:\\/\|\?\*]", "_", name)
        return name or "anexo"

    def _read_attachment_bytes(self, att) -> bytes | None:
        size = None
        for getter in ("size", "get_size", "data_size", "get_data_size"):
            try:
                v = getattr(att, getter)
                v = v() if callable(v) else v
                if isinstance(v, int) and v > 0:
                    size = v
                    break
            except Exception:
                continue
        # 1) read_buffer(size)
        try:
            rb = getattr(att, "read_buffer", None)
            if callable(rb) and size:
                data = rb(size)
                if data:
                    return data
        except Exception:
            pass
        # 2) get_data()
        try:
            gd = getattr(att, "get_data", None)
            if callable(gd):
                data = gd()
                if data:
                    return data
        except Exception:
            pass
        # 3) read()
        try:
            r = getattr(att, "read", None)
            if callable(r):
                data = r(size) if size else r()
                if data:
                    return data
        except Exception:
            pass
        return None

    def _sniff_mime(self, name: str, data: bytes | None) -> str:
        # Prefer header/extension; if puremagic disponível e temos bytes, melhorar detecção
        guessed, _ = mimetypes.guess_type(name)
        if puremagic and data:
            try:
                res = puremagic.from_string(data, mime=True)
                if isinstance(res, str):
                    return res
                if isinstance(res, list) and res:
                    return res[0]
            except Exception:
                pass
        return guessed or "application/octet-stream"

    def get_attachments(self, msg_id: str) -> List[str]:
        msg = self._resolve_message(msg_id)
        display: List[str] = []
        try:
            ac = msg.number_of_attachments
        except Exception:
            ac = getattr(msg, "get_number_of_attachments", lambda: 0)()
        for i in range(ac or 0):
            try:
                att = msg.get_attachment(i)
            except Exception:
                continue
            if self._is_embedded_message(att):
                display.append(f"mensagem incorporada {i} (message/rfc822)")
                continue
            name = self._get_attr(att, ("long_filename", "get_long_filename", "filename", "get_filename"), default=f"anexo_{i}")
            name = self._sanitize_filename(name)
            data = self._read_attachment_bytes(att)
            mime = self._get_attr(att, ("mime_type", "get_mime_type", "mime_tag", "get_mime_tag", "content_type", "get_content_type"), default="")
            mime = mime or self._sniff_mime(name, data)
            display.append(f"{name} ({mime})")
        return display

    def save_attachments(self, msg_id: str, output_dir: str) -> List[str]:
        msg = self._resolve_message(msg_id)
        saved: List[str] = []
        os.makedirs(output_dir, exist_ok=True)
        try:
            ac = msg.number_of_attachments
        except Exception:
            ac = getattr(msg, "get_number_of_attachments", lambda: 0)()
        for i in range(ac or 0):
            try:
                att = msg.get_attachment(i)
            except Exception:
                continue
            if self._is_embedded_message(att):
                continue
            name = self._get_attr(att, ("long_filename", "get_long_filename", "filename", "get_filename"), default=f"anexo_{i}")
            name = self._sanitize_filename(name)
            data = self._read_attachment_bytes(att)
            if not data:
                continue
            out_path = os.path.join(output_dir, name)
            base, ext = os.path.splitext(out_path)
            k = 1
            while os.path.exists(out_path):
                out_path = f"{base} ({k}){ext}"
                k += 1
            with open(out_path, "wb") as f:
                f.write(data)
            saved.append(out_path)
        return saved

    # Helpers
    def _get_attr(self, obj, names: Tuple[str, ...], default: str = "") -> str:
        for n in names:
            try:
                v = getattr(obj, n)
                if callable(v):
                    v = v()
                if v is None:
                    continue
                if isinstance(v, bytes):
                    try:
                        return v.decode("utf-8", errors="replace")
                    except Exception:
                        return default
                return str(v)
            except Exception:
                continue
        return default

    def _normalize_text(self, text: str) -> str:
        return (text or "").replace("\r\n", "\n").replace("\r", "\n")

    def _html_to_text(self, html: str) -> str:
        try:
            import html2text  # type: ignore

            conv = html2text.HTML2Text()
            conv.ignore_links = False
            conv.ignore_images = True
            conv.body_width = 0
            text = conv.handle(html)
            return self._normalize_text(text)
        except Exception:
            return self._normalize_text(html)

    def _to_model_preview(self, msg) -> PstEmail:
        subject = self._get_attr(msg, ("subject", "get_subject"))
        sender = self._get_attr(msg, ("sender_name", "get_sender_name", "sender_email_address", "get_sender_email_address"))
        date = self._get_attr(msg, ("client_submit_time", "get_client_submit_time"))
        return PstEmail(
            id="",
            subject=subject,
            sender=sender,
            to="",
            cc="",
            date=str(date) if date else None,
            body_text=None,
            body_html=None,
            attachments=[],
        )

    def _to_model_full(self, msg, msg_id: str) -> PstEmail:
        subject = self._get_attr(msg, ("subject", "get_subject"))
        sender = self._get_attr(msg, ("sender_name", "get_sender_name", "sender_email_address", "get_sender_email_address"))
        to = self._get_attr(msg, ("display_to", "get_display_to"))
        cc = self._get_attr(msg, ("display_cc", "get_display_cc"))
        date = self._get_attr(msg, ("client_submit_time", "get_client_submit_time"))
        body_text = self._get_attr(msg, ("plain_text_body", "get_plain_text_body"))
        body_html = self._get_attr(msg, ("html_body", "get_html_body"))
        if not body_text and body_html:
            body_text = self._html_to_text(body_html)
        else:
            body_text = self._normalize_text(body_text)
            if body_html:
                body_html = self._normalize_text(body_html)
        names = self.get_attachments(msg_id)
        return PstEmail(
            id=msg_id,
            subject=subject,
            sender=sender,
            to=to,
            cc=cc,
            date=str(date) if date else None,
            body_text=body_text or None,
            body_html=body_html or None,
            attachments=names,
        )
