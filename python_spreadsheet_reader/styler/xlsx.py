from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection, Color
from openpyxl.cell.cell import Cell, MergedCell
from typing import Self, Optional, Literal
from .utils import VerticalAlignment, PatternFills, Underline


type CellOrMegedCell = Cell | MergedCell


class XLSXCellStyler:
    def __init__(self, cell: CellOrMegedCell):
        self.cell = cell

    def font(
        self,
        name: Optional[str] = None,
        size: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[Underline] = None,
        strike: Optional[bool] = None,
        color: Optional[str] = None,
        vert_align: Optional[VerticalAlignment] = None,
    ) -> Self:
        """Applies font styles to the cell.

        Args:
            name: The name of the font family (e.g., "Arial"). Defaults to None.
            size: The font size in points. Defaults to None.
            bold: True to apply bold style. Defaults to None.
            italic: True to apply italic style. Defaults to None.
            underline: The underline style to apply. Defaults to None.
            strike: True to apply strikethrough style. Defaults to None.
            color: The RGB hex color code (e.g., "#FF0000"). Defaults to None.
            vert_align: The vertical alignment style. Defaults to None.

        Returns:
            Self: The styler instance for method chaining.
        """
        current = self.cell.font
        self.cell.font = Font(
            name=name or current.name,
            size=size or current.size,
            bold=bold or current.bold,
            italic=italic or current.italic,
            underline=underline.value if underline else current.underline,
            strike=strike or current.strike,
            color=Color(rgb=color.replace("#", "")) if color else current.color,
            vertAlign=vert_align.value if vert_align else current.vertAlign,
        )
        return self

    def fill(
            self,
            color: str = "#000000",
            fill_type: PatternFills = PatternFills.SOLID,
    ) -> Self:
        """Apply a fill color to the cell.

        Args:
            color: The fill color as a hex string (e.g., "#FF0000").
            fill_type: The pattern fill style to apply.

        Returns:
            Self: The styler instance for method chaining.
        """
        c = color.replace("#", "")
        self.cell.fill = PatternFill(
            start_color=c,
            end_color=c,
            fill_type=fill_type.value
        )
        return self

    def border(
        self,
        sides=("left", "right", "top", "bottom"),
        style="thin",
        color="000000",
    ) -> Self:
        """
        Apply a border style to the given sides only, leaving other sides
        (e.g. from a previous .border() call) untouched.
        """
        side = Side(border_style=style, color=color)
        current = self.cell.border
        sides = set(sides)
        self.cell.border = Border(
            left=side if "left" in sides else current.left,
            right=side if "right" in sides else current.right,
            top=side if "top" in sides else current.top,
            bottom=side if "bottom" in sides else current.bottom,
            diagonal=current.diagonal,
            diagonal_direction=current.diagonal_direction,
            outline=current.outline,
        )
        return self

    def alignment(
        self,
        horizontal=None,
        vertical=None,
        wrap_text=None,
        shrink_to_fit=None,
        text_rotation=None,
        indent=None,
    ) -> Self:
        current = self.cell.alignment
        self.cell.alignment = Alignment(
            horizontal=horizontal if horizontal is not None else current.horizontal,
            vertical=vertical if vertical is not None else current.vertical,
            wrap_text=wrap_text if wrap_text is not None else current.wrap_text,
            shrink_to_fit=shrink_to_fit if shrink_to_fit is not None else current.shrink_to_fit,
            text_rotation=text_rotation if text_rotation is not None else current.text_rotation,
            indent=indent if indent is not None else current.indent,
        )
        return self

    def number_format(self, fmt: str) -> Self:
        self.cell.number_format = fmt
        return self

    def protection(self, locked=None, hidden=None) -> Self:
        current = self.cell.protection
        self.cell.protection = Protection(
            locked=locked if locked is not None else current.locked,
            hidden=hidden if hidden is not None else current.hidden,
        )
        return self

