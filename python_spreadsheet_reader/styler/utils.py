from enum import StrEnum
from pydantic import BaseModel, Field, PositiveInt


# * Type Utilities

class VerticalAlignment(StrEnum):
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"
    JUSTIFY = "justify"
    DISTRIBUTED = "distributed"

class HorizontalAlignment(StrEnum):
    GENERAL = "general"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    FILL = "fill"
    JUSTIFY = "justify"
    CENTERCONTINUOUS = "centerContinuous"
    DISTRIBUTED = "distributed"

class FontVerticalAlignment(StrEnum):
    TOP = "superscript"
    CENTER = "baseline"
    BOTTOM = "subscript"

class PatternFills(StrEnum):
    TRANSPARENT = 'none'
    SOLID = 'solid'
    DARKDOWN = 'darkDown'
    DARKGRAY = 'darkGray'
    DARKGRID = 'darkGrid'
    DARKHORIZONTAL = 'darkHorizontal'
    DARKTRELLIS = 'darkTrellis'
    DARKUP = 'darkUp'
    DARKVERTICAL = 'darkVertical'
    GRAY0625 = 'gray0625'
    GRAY125 = 'gray125'
    LIGHTDOWN = 'lightDown'
    LIGHTGRAY = 'lightGray'
    LIGHTGRID = 'lightGrid'
    LIGHTHORIZONTAL = 'lightHorizontal'
    LIGHTTRELLIS = 'lightTrellis'
    LIGHTUP = 'lightUp'
    LIGHTVERTICAL = 'lightVertical'
    MEDIUMGRAY = 'mediumGray'


class Underline(StrEnum):
    SINGLE = "single"
    DOUBLE = "double"
    SINGLEACCOUNTING = "singleAccounting"
    DOUBLEACCOUNTING = "doubleAccounting"


class Positions(StrEnum):
    LEFT = "left"
    RIGHT = "right"
    TOP = "top"
    BOTTOM = "bottom"

    @classmethod
    def all(cls) -> tuple[Positions, ...]:
        return tuple([v for v in cls])

class BorderStyles(StrEnum):
    DASHDOT = 'dashDot'
    DASHDOTDOT = 'dashDotDot'
    DASHED = 'dashed'
    DOTTED = 'dotted'
    DOUBLE = 'double'
    HAIR = 'hair'
    MEDIUM = 'medium'
    MEDIUMDASHDOT = 'mediumDashDot'
    MEDIUMDASHDOTDOT = 'mediumDashDotDot'
    MEDIUMDASHED = 'mediumDashed'
    SLANTDASHDOT = 'slantDashDot'
    THICK = 'thick'
    THIN = 'thin'

class TextRotationRangeValidator(BaseModel):
    value: PositiveInt = Field(..., ge=0, le=180)

class IndentRangeValidator(BaseModel):
    value: int = Field(..., ge=-255, le=255)