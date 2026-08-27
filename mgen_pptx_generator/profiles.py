from __future__ import annotations

from typing import Any, Literal

from .atelier import ATELIER_ROOTS
from .validation import REQUIRED_ROOTS as CONFERENCE_ROOTS

SupportType = Literal["conference", "atelier"]


def detect_support_type(data: dict[str, Any]) -> SupportType:
    roots = set(data.keys())
    if roots == CONFERENCE_ROOTS or (
        bool(roots) and roots < CONFERENCE_ROOTS
    ):
        return "conference"
    if roots == ATELIER_ROOTS or (
        bool(roots) and roots < ATELIER_ROOTS
    ):
        return "atelier"
    raise ValueError(
        "Format YAML non reconnu. Racines attendues: schéma sémantique Conférence "
        "ou schéma technique Atelier."
    )
