"""Font style definitions."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class Font:
    """Represents a cell font style."""

    name: str | None = "Calibri"
    size: float | None = 11.0
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strike: bool = False
    color: str | None = None
    scheme: str | None = None
    family: int | None = None
    vertAlign: str | None = None
    charset: int | None = None

    def copy(self, **kwargs: Any) -> Font:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {
            "name": self.name,
            "size": self.size,
            "bold": self.bold,
            "italic": self.italic,
            "underline": self.underline,
            "strike": self.strike,
            "color": self.color,
            "scheme": self.scheme,
            "family": self.family,
            "vertAlign": self.vertAlign,
            "charset": self.charset,
        }
        values.update(kwargs)
        return Font(**values)


DEFAULT_FONT = Font()
