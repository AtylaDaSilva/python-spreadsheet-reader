class SpreadsheetReaderException(Exception):
    pass


class NoActiveSpreadsheetException(SpreadsheetReaderException):
    pass


class SpreadsheetIsLockedException(Exception):
    pass
