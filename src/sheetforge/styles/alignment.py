"""Alignment style definitions."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class Alignment:
    """Represents cell alignment settings."""

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int = 0
    text_rotation: int = 0

    def copy(self, **kwargs: Any) -> Alignment:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {
            "horizontal": self.horizontal,
            "vertical": self.vertical,
            "wrap_text": self.wrap_text,
            "shrink_to_fit": self.shrink_to_fit,
            "indent": self.indent,
            "text_rotation": self.text_rotation,
        }
        values.update(kwargs)
        return Alignment(**values)


DEFAULT_ALIGNMENT = Alignment()
