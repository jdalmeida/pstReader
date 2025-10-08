"""
@author João Gbriel de Almeida
"""

from email.message import EmailMessage
from typing import Optional
from src.models import PstEmail


def build_eml(msg: PstEmail) -> str:
    em = EmailMessage()
    em["Subject"] = msg.subject or ""
    em["From"] = msg.sender or ""
    em["To"] = msg.to or ""
    if msg.cc:
        em["Cc"] = msg.cc
    if msg.date:
        em["Date"] = msg.date

    # Preferir HTML se presente; incluir alternativa texto simples
    if msg.body_html and msg.body_text:
        em.set_content(msg.body_text)
        em.add_alternative(msg.body_html, subtype="html")
    elif msg.body_html:
        em.add_alternative(msg.body_html, subtype="html")
    else:
        em.set_content(msg.body_text or "")

    return em.as_string()


def build_txt(msg: PstEmail) -> str:
    headers = [
        f"Assunto: {msg.subject or ''}",
        f"De: {msg.sender or ''}",
        f"Para: {msg.to or ''}",
        f"Cc: {msg.cc or ''}",
        f"Data: {msg.date or ''}",
        "",
    ]
    return "\n".join(headers) + (msg.body_text or "")
