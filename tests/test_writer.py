import os
from tempfile import TemporaryDirectory
from typing import Tuple

import pytest

from excel_writer.writer import ExcelWriter


class TestExcelWriter:
    def setup_method(self):
        self.writer = ExcelWriter(default_sheet_name="original_sheet")

    def test_writer_class(self):
        assert self.writer.worksheets == ("original_sheet",)

    def test_rename_sheet(self):
        self.writer.rename_sheet(0, new_sheet_name="new_sheet_name")
        assert self.writer.worksheets == ("new_sheet_name",)

    def test_create_sheet(self):
        self.writer.create_sheet("new_sheet")
        assert self.writer.worksheets == ("original_sheet", "new_sheet")

    @pytest.mark.parametrize(
        "sheet_names, expected_order",
        [
            (
                ("sheet_1", "sheet_2", "sheet_3"),
                ("original_sheet", "sheet_1", "sheet_2", "sheet_3"),
            ),
            (
                ("sheet_5", "sheet_4", "sheet_3"),
                ("original_sheet", "sheet_5", "sheet_4", "sheet_3"),
            ),
        ],
    )
    def test_worksheet_property_ordered(
        self, sheet_names: Tuple[str, ...], expected_order: Tuple[str, ...]
    ):
        for sheet in sheet_names:
            self.writer.create_sheet(sheet)
        assert self.writer.worksheets == expected_order

    def test_save_workbook(self):
        with TemporaryDirectory() as tmpdir:
            self.writer.save_workbook(filepath=tmpdir, filename="test.xlsx")
            assert os.path.isfile(os.path.join(tmpdir, "test.xlsx"))

    def test_set_current_sheet_by_name(self):
        self.writer.create_sheet("new_sheet")
        self.writer.set_active_sheet("new_sheet")
        assert self.writer.active_sheet.title == "new_sheet"

    def test_set_current_sheet_by_index(self):
        self.writer.create_sheet("new_sheet_1")
        self.writer.create_sheet("new_sheet_2")
        self.writer.set_active_sheet(2)
        assert self.writer.active_sheet.title == "new_sheet_2"

    def test_cell_using_notation(self):
        cell = self.writer.cell(0, "A1", set_value="test_value")
        assert cell.value == "test_value"

    def test_cell_using_row_col(self):
        cell = self.writer.cell(0, (1, 2), set_value="test_value")
        assert cell.value == "test_value"
