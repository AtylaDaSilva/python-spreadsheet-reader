from enum import StrEnum


# * Type Enums

class VerticalAlignment(StrEnum):
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
