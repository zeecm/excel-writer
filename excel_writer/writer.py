import os
from typing import Optional, Protocol, Set, Union, overload

from openpyxl import Workbook


class Writer(Protocol):
    def write_cell(self, cell_notation: str) -> None:
        ...


class ExcelWriter:
    def __init__(
        self, default_sheet_name: Optional[str] = None, iso_dates: bool = False
    ):
        self._workbook = Workbook(iso_dates=iso_dates)
        self._current_sheet = self._workbook.active
        if default_sheet_name is not None:
            self._current_sheet.title = default_sheet_name

    @property
    def worksheets(self) -> Set[str]:
        return set(self._workbook.sheetnames)

    def create_sheet(self, sheet_name: str, position: Optional[str] = None) -> None:
        self._workbook.create_sheet(sheet_name, position)

    def rename_sheet(self, sheet: Union[str, int], new_sheet_name: str) -> None:
        self._set_current_sheet(sheet)
        self._current_sheet.title = new_sheet_name

    def save_workbook(self, filepath: str, filename: str) -> None:
        full_filepath = os.path.join(filepath, filename)
        self._workbook.save(full_filepath)

    @overload
    def set_cell_value(
        self, sheet: Union[str, int], row: int, column: int, value: str
    ) -> None:
        ...

    @overload
    def set_cell_value(self, sheet: Union[str, int], cell: str, value: str) -> None:
        ...

    def set_cell_value(self, sheet: Union[str, int], value: str, **kwargs) -> None:
        if ["row", "column"] in kwargs:
            self._set_current_sheet(sheet)
            self._current_sheet.cell(
                row=kwargs["row"], column=kwargs["column"], value=value
            )
        if "cell" in kwargs:
            self._set_current_sheet(sheet)
            self._current_sheet[kwargs["cell"]] = value

    def _set_current_sheet(self, sheet: Optional[Union[str, int]] = None) -> None:
        """Sets current sheet, if sheet is not provided, defaults to first sheet"""
        if sheet is None:
            self._current_sheet = self._workbook.worksheets[0]
        if isinstance(sheet, str):
            self._current_sheet = self._workbook[sheet]
        if isinstance(sheet, int):
            self._current_sheet = self._workbook.worksheets[sheet]
