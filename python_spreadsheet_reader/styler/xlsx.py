from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection, Color
from openpyxl.cell.cell import Cell, MergedCell
from typing import Self, Optional
from .utils import (
    HorizontalAlignment,
    VerticalAlignment,
    FontVerticalAlignment,
    PatternFills,
    Underline,
    Positions,
    BorderStyles,
    TextRotationRangeValidator,
    IndentRangeValidator,
)
from pydantic import ValidationError


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
        vert_align: Optional[FontVerticalAlignment] = None,
    ) -> Self:
        """Applies font styles to the cell.

        Args:
            name: The name of the font family (e.g., "Arial"). Defaults to None.
            size: The font size in points. Defaults to None.
            bold: True to apply bold style. Defaults to None.
            italic: True to apply italic style. Defaults to None.
            underline: The underline style to apply. Defaults to None.
            strike: True to apply strikethrough style. Defaults to None.
            color: The RGB hex color code (e.g., "#000000"). Defaults to None.
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
            color: The fill color as a hex string (e.g., "#000000").
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
        sides: tuple[Positions, ...] = Positions.all(),
        style: BorderStyles = BorderStyles.THIN,
        color: str = "#000000",
    ) -> Self:
        """Apply a border style to the cell.

        Args:
            sides: A tuple of border positions to apply the style to (for example,
                left, right, top, bottom). Defaults to all positions.
            style: The border style to use. Defaults to :pydata:`BorderStyles.THIN`.
            color: The border color as a hex string (for example, "#000000").
                The leading hash is optional.

        Returns:
            Self: The styler instance for method chaining.
        """
        current = self.cell.border
        border = Border(
            diagonal=current.diagonal,
            diagonal_direction=current.diagonal_direction,
            outline=current.outline,
        )
        for s in sides:
            setattr(border, s.name.lower(), Side(border_style=style, color=color.replace("#", "")))
        self.cell.border = border
        return self

    def alignment(
        self,
        horizontal: Optional[HorizontalAlignment] = None,
        vertical: Optional[VerticalAlignment] = None,
        wrap_text: Optional[bool] = None,
        shrink_to_fit: Optional[bool] = None,
        text_rotation: Optional[int] = None,
        indent=None,
    ) -> Self:
        """Apply alignment settings to the cell.

        Args:
            horizontal: The horizontal alignment to apply.
            vertical: The vertical alignment to apply.
            wrap_text: Whether to wrap text within the cell.
            shrink_to_fit: Whether to shrink text to fit the cell.
            text_rotation: The text rotation angle, between 0 and 180.
            indent: The indentation level, between -255 and 255.

        Returns:
            Self: The styler instance for method chaining.

        Raises:
            ValueError: If text_rotation is outside the allowed range.
            ValueError: If indent is outside the allowed range.
        """
        current = self.cell.alignment
        try:
            tr = (
                TextRotationRangeValidator(value=text_rotation).value
                if text_rotation is not None else current.text_rotation
            )
        except ValidationError:
            raise ValueError(f"Text rotation must be between 0 and 180, got: {text_rotation}")

        try:
            ind = (
                IndentRangeValidator(value=indent).value
                if indent is not None else current.indent
            )
        except ValidationError:
            raise ValueError(f"Indent must be between -255 and 255, got: {indent}")

        self.cell.alignment = Alignment(
            horizontal=horizontal.value if horizontal else current.horizontal,
            vertical=vertical.value if vertical else current.vertical,
            wrap_text=wrap_text or current.wrap_text,
            shrink_to_fit=shrink_to_fit or current.shrink_to_fit,
            text_rotation=tr,
            indent=ind
        )
        return self

    def number_format(self, fmt: str = "General") -> Self:
        """Set the number format for the cell.

        Args:
            fmt:
                The number format string to apply.
                Can either be a build-in literal (like 'General') or a custom format string, e.g.:

                ``"0.00"``, ``"#,##0"``, ``"#,##0.00"``, ``"0%"``, ``"0.00%"``,
                ``"mm-dd-yy"``, "h:mm:ss AM/PM", ``"h:mm"``, ``"h:mm:ss"``, ``"m/d/yy h:mm"``.

        Returns:
            Self: The styler instance for method chaining.
        """
        self.cell.number_format = fmt
        return self
