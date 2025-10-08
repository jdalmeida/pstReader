"""
@author João Gbriel de Almeida
"""

from typing import List
import shutil

from src.models import PstFolder, PstEmail


class BaseAdapter:
    def open(self, path: str) -> None:  # pragma: no cover
        raise NotImplementedError

    def get_root_folders(self) -> List[PstFolder]:  # pragma: no cover
        raise NotImplementedError

    def list_messages(self, folder_id: str) -> List[PstEmail]:  # pragma: no cover
        raise NotImplementedError

    def get_message(self, msg_id: str) -> PstEmail:  # pragma: no cover
        raise NotImplementedError

    def export_eml(self, msg_id: str, out_path: str) -> None:  # pragma: no cover
        raise NotImplementedError

    def get_attachments(self, msg_id: str) -> List[str]:  # pragma: no cover
        raise NotImplementedError

    def save_attachments(self, msg_id: str, output_dir: str) -> List[str]:  # returns saved file paths
        raise NotImplementedError


class PstReader:
    def __init__(self) -> None:
        self.adapter: BaseAdapter | None = None

    def open(self, path: str) -> None:
        # Prefer pypff
        try:
            from src.adapters.pypff_adapter import PypffAdapter  # lazy import
        except ImportError:
            self.adapter = None
        else:
            try:
                adapter = PypffAdapter()
                adapter.open(path)
                self.adapter = adapter
                return
            except Exception as exc:
                # Não fazer fallback silencioso: informe erro real de abertura
                raise RuntimeError(f"Falha ao abrir PST com pypff: {exc}") from exc

        # Fallback: readpst disponível no PATH?
        if shutil.which("readpst"):
            from src.adapters.readpst_adapter import ReadPstAdapter

            adapter = ReadPstAdapter()
            adapter.open(path)
            self.adapter = adapter
            return

        raise RuntimeError(
            "Nenhum adaptador disponível: instale pypff/libpff ou disponibilize readpst no PATH."
        )

    def _require(self):
        if not self.adapter:
            raise RuntimeError("PST não aberto")
        return self.adapter

    def get_root_folders(self) -> List[PstFolder]:
        return self._require().get_root_folders()

    def list_messages(self, folder_id: str) -> List[PstEmail]:
        return self._require().list_messages(folder_id)

    def get_message(self, msg_id: str) -> PstEmail:
        return self._require().get_message(msg_id)

    def export_eml(self, msg_id: str, out_path: str) -> None:
        return self._require().export_eml(msg_id, out_path)

    def get_attachments(self, msg_id: str) -> List[str]:
        return self._require().get_attachments(msg_id)

    def save_attachments(self, msg_id: str, output_dir: str) -> List[str]:
        return self._require().save_attachments(msg_id, output_dir)
