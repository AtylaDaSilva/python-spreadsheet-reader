from python_spreadsheet_reader.readers.exceptions import SpreadsheetIsLockedException
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.drawing.image import Image
from pathlib import Path
from typing import Any
import os
import re


class XLSXReader:
    def __init__(self, workbook_path: str | Path) -> None:
        self.workbook_path = (
            Path(workbook_path)
            if isinstance(workbook_path, str)
            else workbook_path
        )
        self._sheet_data: dict[int, dict] = {}  # Local cache

    # * ---------------------------------------------------------------------------
    # *                               Public API
    # * ---------------------------------------------------------------------------

    def read_sheet(
            self, sheet_name: str | None = None,
            read_only: bool = True,
            cell_values_only: bool = False,
            return_cell_coords: bool = True,
            preserve_formulas: bool = True,
            read_locked: bool = False
    ) -> dict[int, dict]:
        """
        Returns the data from the spreadsheet located at *self.workbook_path*.
        Args:
            sheet_name:
                Name of the sheet to read data from. If not specified, returns the data from the active sheet.
            read_only:
                Opens the workbook in an optimized for reading mode, but content can't be edited. Defaults to True.
            cell_values_only:
                If True, returns cell values (str, int, datetime, etc.) instead of Cell objects. Defaults to False.
            return_cell_coords:
                If True, returns cell coordinates (A1, B2, C3, ect.) instead of column numbers.
                Does not affect row numbers. Defaults to True.
            preserve_formulas:
                If True, returns cell formulae instead of cached values. Defaults to True.
            read_locked:
                If True, allows the reading of locked (currently opened) spreadsheets. Defaults to False.

        Returns: A dict of sheet rows, each key representing the row number (1-based) and each value a nested dict.
        The nested dicts represents cells, with cell coordinates (or column numbers) as keys and cell values (or objects) as values.

        Examples:
            >>>reader = XLSXReader(workbook_path="path/to/workbook.xlsx")
            >>>sheet_data = reader.read_sheet("Sheet2", cell_values_only=True)

            >>>print(sheet_name)

            # Outputs

            {
                1: {"A1", "ID", "A2": "Title", "A3": "Genre"},

                2: {"B1", 123, "B2": "Alien", "A3": "Science Fiction, Horror"},

                2: {"C1", 124, "C2": "Predator", "A3": "Science Fiction, Action"},

                ...
            }

        """
        # Validations
        if not self.workbook_path.exists():
            raise FileNotFoundError(
                f'Spreadsheet not found: "{self.workbook_path}"'.replace("\\", "/")  # Formatar \\ no caminho do arquivo
            )
        elif not read_locked and self.is_spreadsheet_locked():
            raise SpreadsheetIsLockedException(f'Cannot read locked spreadsheet: "{self.workbook_path}"')

        match self.workbook_path.suffix.lower():
            case ".xlsx":
                # Open active workbook (or specific, if provided)
                wb = openpyxl.load_workbook(
                    self.workbook_path,
                    read_only=read_only,
                    data_only=(not preserve_formulas)
                )
                if sheet_name:
                    ws = wb[sheet_name]
                else:
                    ws = wb.active

                # Iterate through rows/cols to get cell values
                rows = {}
                for i, row in enumerate(ws.iter_rows()):
                    r = {}
                    for cell in row:
                        if isinstance(cell, openpyxl.cell.read_only.EmptyCell):  # Ignore empty cells
                            continue
                        r[cell.coordinate if return_cell_coords else cell.column] = cell.value if cell_values_only else cell
                    if len(r) > 0:  # Ignore empty rows
                        rows[i + 1] = r
                self._sheet_data = rows
                wb.close()
                return self._sheet_data
            case _:
                raise ValueError(f'Unsupported file type: "{self.workbook_path.suffix}".')

    def get_cell(self, row: int, col: int) -> Cell | MergedCell:
        raise NotImplementedError
        ws = self._planilha_ativa()
        return ws.cell(row=row, column=col)

    def set_cell_value(self, linha: int, coluna: int, valor: Any) -> None:
        raise NotImplementedError
        self.get_celula(linha=linha, coluna=coluna).value = valor

    def adicionar_imagem(
        self, caminho_imagem: str, celula: str, altura: int, largura: int
    ):
        raise NotImplementedError
        imagem = Image(caminho_imagem)
        imagem.height = altura
        imagem.width = largura
        ws = self._planilha_ativa()
        ws.add_image(imagem, celula)

    def salvar_planilha(self, caminho: str | Path | None = None) -> Path:
        raise NotImplementedError
        wb = self._abrir_workbook()
        p: Path = self.caminho_planilha if not caminho else Path(caminho)
        if not p.parent.is_dir():
            p.parent.mkdir(parents=True, exist_ok=True)
        wb.save(p)
        return p

    def is_spreadsheet_locked(self) -> bool:
        """
        Return True if the spreadsheet at filepath is currently locked by
        another process, False otherwise.

        Checks for lock files created by Excel (~$<filename>) and LibreOffice
        (.~lock.<filename>#) in the same directory as the target file.

        Parameters
        ----------
        filepath:
            Path to the target spreadsheet.

        Raises
        ------
        FileNotFoundError
            If filepath does not exist.
        """
        if not self.workbook_path.exists():
            raise FileNotFoundError(f'Could not find spreadsheet: "{self.workbook_path}"')

        return self._is_xlsx_locked() or self._is_ods_locked()

    # * ---------------------------------------------------------------------------
    # *                                Internal API
    # * ---------------------------------------------------------------------------

    def _abrir_workbook(self, *args, **kwargs) -> openpyxl.Workbook:
        raise NotImplementedError
        if self.is_spreadsheet_locked():
            raise SpreadsheetIsLockedException(self.caminho_planilha)
        if self._workbook is not None:
            return self._workbook
        wb = openpyxl.load_workbook(self.caminho_planilha, *args, **kwargs)
        self._workbook = wb
        return wb

    def _fechar_workbook(self):
        raise NotImplementedError
        if self._workbook:
            self._workbook.close()
            self._workbook = None
            self._sheet_data = {}

    def _planilha_ativa(self):
        raise NotImplementedError
        active_sheet = self._abrir_workbook().active
        if active_sheet is None:
            raise NoActiveSheetException(
                f"Não foi possível localizar a planilha ativa no excel {self.caminho_planilha}"
            )
        return active_sheet

    def _is_xlsx_locked(self) -> bool:
        """
        Return True if an Excel lock file (~$<filename>) exists in the same
        directory as filepath, False otherwise.

        Excel creates a temporary lock file prefixed with "~$" when a workbook
        is open. Its presence indicates the file is currently in use.
        """
        return bool(re.search(r"~\$" + re.escape(self.workbook_path.name), " ".join(
            f.name for f in self.workbook_path.parent.iterdir()
        )))

    def _is_ods_locked(self) -> bool:
        """
        Return True if a LibreOffice lock file (.~lock.<filename>#) exists in
        the same directory as filepath, False otherwise.

        LibreOffice creates a lock file with the pattern .~lock.<filename># when
        a document is open. Its presence indicates the file is currently in use.
        """
        return bool(re.search(r"\.~lock\." + re.escape(self.workbook_path.name) + r"#", " ".join(
            f.name for f in self.workbook_path.parent.iterdir()
        )))
