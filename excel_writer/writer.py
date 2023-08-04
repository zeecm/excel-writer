import os
from typing import Optional, Protocol, Tuple, Union, overload

from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


class Writer(Protocol):
    def save_workbook(self, filepath: str, filename: str) -> None:
        ...


class ExcelWriter:
    def __init__(
        self,
        existing_workbook: str = "",
        default_sheet_name: str = "",
        iso_dates: bool = False,
    ):
        self._workbook = self._initialize_workbook(
            existing_workbook, default_sheet_name, iso_dates
        )
        self.active_sheet = self._get_active_sheet(self._workbook)

    def _initialize_workbook(
        self,
        existing_workbook: str = "",
        default_sheet_name: str = "",
        iso_dates: bool = False,
    ) -> Workbook:
        if existing_workbook:
            return self._load_existing_workbook(existing_workbook)
        return self._create_new_workbook(default_sheet_name, iso_dates)

    def _load_existing_workbook(self, filepath: str) -> Workbook:
        return load_workbook(filepath)

    def _create_new_workbook(
        self, default_sheet_name: str = "", iso_dates: bool = False
    ) -> Workbook:
        workbook = Workbook(iso_dates=iso_dates)
        active_sheet = self._get_active_sheet(workbook)
        if default_sheet_name:
            active_sheet.title = default_sheet_name
        return workbook

    def _get_active_sheet(self, workbook: Workbook) -> Worksheet:
        if not workbook.active:
            workbook.active = self._get_worksheet(0)
        # correct typing, worksheet is subclass of _WorkbookChild
        return workbook.active  # type: ignore

    @property
    def worksheets(self) -> Tuple[str, ...]:
        return tuple(self._workbook.sheetnames)

    def create_sheet(self, sheet_name: str, position: Optional[int] = None) -> None:
        self._workbook.create_sheet(sheet_name, position)
        self.set_current_sheet(sheet_name)

    def rename_sheet(self, sheet: Union[str, int], new_sheet_name: str) -> None:
        sheet_obj = self._get_worksheet(sheet)
        sheet_obj.title = new_sheet_name

    def save_workbook(self, filepath: str, filename: str) -> None:
        full_filepath = os.path.join(filepath, filename)
        self._workbook.save(full_filepath)

    @overload
    def cell(
        self,
        sheet: Union[str, int],
        cell_id: Tuple[int, int],
        set_value: Optional[str] = None,
    ) -> Cell:
        ...

    @overload
    def cell(
        self,
        sheet: Union[str, int],
        cell_id: str,
        set_value: Optional[str] = None,
    ) -> Cell:
        ...

    def cell(
        self,
        sheet: Union[str, int],
        cell_id: Union[Tuple[int, int], str],
        set_value: Optional[str] = None,
    ) -> Cell:
        sheet_object = self._get_worksheet(sheet)

        if isinstance(cell_id, tuple):
            cell = self._get_cell_by_row_col(sheet_object, row_col=cell_id)
        elif isinstance(cell_id, str):
            cell = self._get_cell_by_notation(sheet_object, cell_notation=cell_id)
        else:
            raise ValueError("one of row and column or cell notation must be specified")
        if set_value:
            cell.value = set_value
        return cell

    def _get_cell_by_row_col(self, sheet: Worksheet, row_col: Tuple[int, int]) -> Cell:
        row, column = row_col
        return sheet.cell(row=row, column=column)

    def _get_cell_by_notation(self, sheet: Worksheet, cell_notation: str) -> Cell:
        return sheet[cell_notation]

    @overload
    def _get_worksheet(self, sheet: str) -> Worksheet:
        ...

    @overload
    def _get_worksheet(self, sheet: int) -> Worksheet:
        ...

    def _get_worksheet(self, sheet: Union[str, int]):
        if isinstance(sheet, str):
            return self._workbook[sheet]
        if isinstance(sheet, int):
            return self._workbook.worksheets[sheet]
        raise ValueError(f"invalid sheet argument {sheet}")

    def set_current_sheet(self, sheet: Optional[Union[str, int]] = None) -> None:
        """Sets current sheet, if sheet is not provided, defaults to first sheet"""
        if sheet is None:
            self.active_sheet = self._workbook.worksheets[0]
            return
        self.active_sheet = self._get_worksheet(sheet)
