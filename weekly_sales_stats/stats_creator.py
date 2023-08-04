from typing import Optional

from excel_writer.writer import ExcelWriter, Writer


class WeeklySalesStatsGenerator:
    def __init__(self, previous_sheet_location: str, writer: Optional[Writer] = None):
        self.writer = writer or ExcelWriter(existing_workbook=previous_sheet_location)

    def create_new_week_tab(self) -> None:
        current_week = self.writer.worksheets[0]
