"""
@author João Gbriel de Almeida
"""

from dataclasses import dataclass
from typing import List, Optional


@dataclass
class PstFolder:
    id: str
    name: str
    children: List["PstFolder"]


@dataclass
class PstEmail:
    id: str
    subject: str
    sender: str
    to: str
    cc: str
    date: Optional[str]
    body_text: Optional[str]
    body_html: Optional[str]
    attachments: List[str]
