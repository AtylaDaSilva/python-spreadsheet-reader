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
