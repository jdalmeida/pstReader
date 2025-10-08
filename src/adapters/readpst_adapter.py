"""
@author João Gbriel de Almeida
"""

import shutil
from typing import List

from src.models import PstEmail, PstFolder


class ReadPstAdapter:
    def __init__(self) -> None:
        self._path: str | None = None
        if not shutil.which("readpst"):
            raise RuntimeError("readpst não encontrado no PATH")

    def open(self, path: str) -> None:
        # Estratégia futura: converter PST para mbox temporário e indexar
        # No MVP, apenas valida a existência e falha com mensagem orientativa
        self._path = path
        raise RuntimeError(
            "Fallback readpst ainda não implementado. Instale pypff/libpff para melhor experiência."
        )

    def get_root_folders(self) -> List[PstFolder]:  # pragma: no cover
        raise NotImplementedError

    def list_messages(self, folder_id: str) -> List[PstEmail]:  # pragma: no cover
        raise NotImplementedError

    def get_message(self, msg_id: str) -> PstEmail:  # pragma: no cover
        raise NotImplementedError

    def export_eml(self, msg_id: str, out_path: str) -> None:  # pragma: no cover
        raise NotImplementedError
