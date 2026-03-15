"""Border style definitions."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class Side:
    """Represents one side of a cell border."""

    style: str | None = None
    color: str | None = None

    def copy(self, **kwargs: Any) -> Side:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {"style": self.style, "color": self.color}
        values.update(kwargs)
        return Side(**values)


@dataclass
class Border:
    """Represents a cell border (all four sides + diagonal)."""

    left: Side | None = None
    right: Side | None = None
    top: Side | None = None
    bottom: Side | None = None
    diagonal: Side | None = None
    outline: bool = True
    diagonalUp: bool = False
    diagonalDown: bool = False

    def copy(self, **kwargs: Any) -> Border:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {
            "left": self.left,
            "right": self.right,
            "top": self.top,
            "bottom": self.bottom,
            "diagonal": self.diagonal,
            "outline": self.outline,
            "diagonalUp": self.diagonalUp,
            "diagonalDown": self.diagonalDown,
        }
        values.update(kwargs)
        return Border(**values)


DEFAULT_BORDER = Border()
