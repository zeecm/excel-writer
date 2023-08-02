import os
from tempfile import TemporaryDirectory

from excel_writer.writer import ExcelWriter


class TestExcelWriter:
    def setup_method(self):
        self.writer = ExcelWriter(default_sheet_name="original_sheet")

    def test_writer_class(self):
        assert self.writer.worksheets == {"original_sheet"}

    def test_create_sheet(self):
        self.writer.create_sheet("new_sheet")
        assert self.writer.worksheets == {"original_sheet", "new_sheet"}

    def test_save_workbook(self):
        with TemporaryDirectory() as tmpdir:
            self.writer.save_workbook(filepath=tmpdir, filename="test.xlsx")
            assert os.path.isfile(os.path.join(tmpdir, "test.xlsx"))

    def test_set_current_sheet_by_name(self):
        self.writer.create_sheet("new_sheet")
        self.writer._set_current_sheet("new_sheet")
        assert self.writer._current_sheet.title == "new_sheet"

    def test_set_current_sheet_by_index(self):
        self.writer.create_sheet("new_sheet_1")
        self.writer.create_sheet("new_sheet_2")
        self.writer._set_current_sheet(2)
        assert self.writer._current_sheet.title == "new_sheet_2"
