from __future__ import annotations

import os
from typing import List, Optional, Protocol, Tuple, Type, Union, overload

from loguru import logger
from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell
from openpyxl.utils import coordinate_to_tuple, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

_CellTypes = Type[Cell]

_CellID = Union[Tuple[int, int], str]


class Writer(Protocol):
    worksheets: Tuple[str, ...]

    def cell(
        self,
        sheet: Union[str, int],
        cell_id: Union[Tuple[int, int], str],
        set_value: Optional[str] = None,
    ) -> _CellTypes:
        ...

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
            logger.info(f"loading existing workbook from {existing_workbook}")
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
            workbook.active = self.get_worksheet(0)
        # correct typing, worksheet is subclass of _WorkbookChild
        return workbook.active  # type: ignore

    @property
    def worksheets(self) -> Tuple[str, ...]:
        return tuple(self._workbook.sheetnames)

    def create_sheet(
        self,
        sheet_name: str,
        position: Optional[int] = None,
        from_worksheet: Optional[Worksheet] = None,
    ) -> None:
        self._workbook.create_sheet(sheet_name, position)
        self.set_active_sheet(sheet_name)
        if from_worksheet is not None:
            new_worksheet = self.get_worksheet(sheet_name)
            self._copy_rows_from_worksheet(from_worksheet, to_worksheet=new_worksheet)

    def _copy_rows_from_worksheet(
        self, from_worksheet: Worksheet, to_worksheet: Worksheet
    ) -> Worksheet:
        for row in from_worksheet.iter_rows(values_only=True):
            to_worksheet.append(row)
        return to_worksheet

    def rename_sheet(self, sheet: Union[str, int], new_sheet_name: str) -> None:
        sheet_obj = self.get_worksheet(sheet)
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
        cell_id: _CellID,
        set_value: Optional[str] = None,
    ) -> Cell:
        sheet_object = self.get_worksheet(sheet)

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
    def get_worksheet(self, sheet: str) -> Worksheet:
        ...

    @overload
    def get_worksheet(self, sheet: int) -> Worksheet:
        ...

    def get_worksheet(self, sheet: Union[str, int]) -> Worksheet:
        if isinstance(sheet, str):
            return self._workbook[sheet]
        if isinstance(sheet, int):
            return self._workbook.worksheets[sheet]
        raise ValueError(f"invalid sheet argument {sheet}")

    def set_active_sheet(self, sheet: Optional[Union[str, int]] = None) -> None:
        """Sets current sheet, if sheet is not provided, defaults to first sheet"""
        if sheet is None:
            self.active_sheet = self._workbook.worksheets[0]
            return
        self.active_sheet = self.get_worksheet(sheet)

    def paste_array(
        self, sheet: Union[str, int], array: List[List[str]], start_cell: _CellID = "A1"
    ) -> None:
        cell_id = start_cell
        if isinstance(cell_id, str):
            cell_id = coordinate_to_tuple(cell_id)

        current_row, current_col = cell_id
        for row in array:
            for value in row:
                notation = self._convert_row_col_to_notation((current_row, current_col))
                self.cell(sheet, notation, set_value=value)
                current_col += 1
            current_row += 1
            current_col = cell_id[1]

    def _convert_row_col_to_notation(self, row_col: Tuple[int, int]) -> str:
        row, col = row_col
        col_letter = get_column_letter(col)
        return f"{col_letter}{row}"

    def copy_worksheet(self, sheet: Union[str, int]) -> Worksheet:
        worksheet = self.get_worksheet(sheet)
        # correct return type, if read only, will raise value error
        return self._workbook.copy_worksheet(worksheet)  # type: ignore


class ExcelWriterContextManager:
    def __init__(
        self,
        existing_workbook: str = "",
        default_sheet_name: str = "",
        iso_dates: bool = False,
    ):
        self._existing_workbook = existing_workbook
        self._default_sheet_name = default_sheet_name
        self._iso_dates = iso_dates
        self.writer = None

    def __enter__(self) -> ExcelWriter:
        self.writer = ExcelWriter(
            self._existing_workbook, self._default_sheet_name, self._iso_dates
        )
        return self.writer

    def __exit__(self, *args, **kwargs):
        if self.writer is not None:
            self.writer._workbook.close()
        self.writer = None
