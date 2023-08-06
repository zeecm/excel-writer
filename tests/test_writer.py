import os
from tempfile import TemporaryDirectory
from typing import List, Tuple

import pytest

from excel_writer.writer import ExcelWriter, ExcelWriterContextManager


class TestExcelWriter:
    def setup_method(self):
        self.writer = ExcelWriter(default_sheet_name="original_sheet")

    @pytest.fixture
    def fixture_test_array(self) -> List[List[str]]:
        return [
            ["this", "is", "a", "test", "array"],
            ["row", "two", "of", "test", "array"],
            ["final", "row", "of", "test", "array"],
        ]

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

    def test_paste_array(self, fixture_test_array: List[List[str]]):
        self.writer.paste_array(0, fixture_test_array, start_cell="B4")
        b4_cell = self.writer.cell(0, "b4")
        assert b4_cell.value == "this"

    def test_copy_worksheet(self, fixture_test_array: List[List[str]]):
        self.writer.paste_array(0, fixture_test_array)
        copied_worksheet = self.writer.copy_worksheet(0)
        assert copied_worksheet.cell(3, 1).value == "final"


def test_excel_writer_context_manager():
    context_manager = None
    with ExcelWriterContextManager(default_sheet_name="test_sheet") as context_manager:
        assert context_manager.worksheets == ("test_sheet",)
