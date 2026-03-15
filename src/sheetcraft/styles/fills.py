"""Fill style definitions."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class PatternFill:
    """Represents a pattern fill style."""

    patternType: str | None = None
    fgColor: str | None = None
    bgColor: str | None = None

    def copy(self, **kwargs: Any) -> PatternFill:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {
            "patternType": self.patternType,
            "fgColor": self.fgColor,
            "bgColor": self.bgColor,
        }
        values.update(kwargs)
        return PatternFill(**values)


@dataclass
class GradientFill:
    """Represents a gradient fill style."""

    type: str = "linear"
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0
    stop: list[str] | None = None

    def copy(self, **kwargs: Any) -> GradientFill:
        """Return a copy with optional overrides."""
        values: dict[str, Any] = {
            "type": self.type,
            "degree": self.degree,
            "left": self.left,
            "right": self.right,
            "top": self.top,
            "bottom": self.bottom,
            "stop": self.stop,
        }
        values.update(kwargs)
        return GradientFill(**values)


FILL_NONE = PatternFill(patternType="none")
FILL_SOLID = PatternFill(patternType="solid")
