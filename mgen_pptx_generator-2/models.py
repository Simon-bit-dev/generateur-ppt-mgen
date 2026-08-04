from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

TextStyle = Literal["plain", "bullet_text", "numbered_text"]


@dataclass(frozen=True)
class PlaceholderValue:
    idx: int
    value: str | list[str]
    style: TextStyle = "plain"
    required: bool = True


@dataclass(frozen=True)
class SlideSpec:
    layout_name: str
    role: str
    placeholders: tuple[PlaceholderValue, ...] = field(default_factory=tuple)
    clone_slide_number: bool = True
