from typing import Optional

from excel_writer.writer import ExcelWriter, Writer


class WeeklySalesStatsGenerator:
    def __init__(self, previous_sheet_location: str, writer: Optional[Writer] = None):
        self.writer = writer or ExcelWriter(existing_workbook=previous_sheet_location)

    def create_new_week_tab(self) -> None:
        latest_tab_name = self.writer.worksheets[0]
        new_tab_name = self._increment_tab_week(latest_tab_name)
        self.writer.create_sheet(new_tab_name)

    def _increment_tab_week(self, tab_name: str, increment: int = 1) -> str:
        previous_week = tab_name[-2:]
        incremented_week = int(previous_week) + increment
        return tab_name[:-2] + str(incremented_week)

    def save_file(self, filepath: str, filename: str) -> None:
        self.writer.save_workbook(filepath, filename)
