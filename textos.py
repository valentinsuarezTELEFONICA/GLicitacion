# -*- coding: utf-8 -*-
"""Utilidades de texto comunes."""
from __future__ import annotations

import unicodedata


def norm(texto: str) -> str:
    """Minúsculas y sin tildes, conservando la longitud (para poder localizar extractos)."""
    return "".join(unicodedata.normalize("NFD", c)[0] for c in str(texto or "")).lower()


def euros(v) -> str:
    return f"{v:,.0f} €".replace(",", ".") if v is not None else "—"
