import os
from typing import Optional, Union

from excel_writer.writer import ExcelWriter, ExcelWriterContextManager, Writer
from weekly_sales_stats.utils import DateRange, calculate_week_dates

TEMPLATE_PATH = os.path.join("templates", "template_weekly_stats.xlsx")


class WeeklySalesStatsGenerator:
    def __init__(self, previous_sheet_location: str, writer: Optional[Writer] = None):
        self.writer = writer or ExcelWriter(existing_workbook=previous_sheet_location)

    def create_new_week_tab(self) -> None:
        new_tab_name = self._get_new_tab_name()
        with ExcelWriterContextManager(TEMPLATE_PATH) as writer:
            template_sheet = writer.copy_worksheet(0)
            self.writer.create_sheet(
                sheet_name=new_tab_name, position=0, from_worksheet=template_sheet
            )
            self.writer.get_worksheet(new_tab_name).sheet_view.showGridLines = False

    def _get_new_tab_name(self) -> str:
        latest_tab_name = self.writer.worksheet_names[0]
        return self._increment_tab_week(latest_tab_name)

    def _increment_tab_week(self, tab_name: str, increment: int = 1) -> str:
        previous_week = tab_name[-2:]
        incremented_week = int(previous_week) + increment
        return tab_name[:-2] + str(incremented_week)

    def _set_data_summary_date(self, sheet: Optional[Union[str, int]] = 0) -> None:
        week_str = self._get_sheet_name(sheet)
        week_num = self._get_week_num_from_week_str(week_str)
        week_dates = calculate_week_dates(week_num)

        formatted_date_range = self._format_date_range(week_dates)
        final_date_str = f"{week_str} {formatted_date_range}"

        self.writer.cell(sheet=sheet, cell_id="D2", set_value=final_date_str)

    def _get_sheet_name(self, sheet: Union[str, int]) -> str:
        return self.writer.get_worksheet(sheet).title

    def _get_week_num_from_week_str(self, week_str: str) -> int:
        return int(week_str[-2:])

    def _format_date_range(self, date_range: DateRange) -> str:
        start_day = date_range.start_date.day
        end_date = date_range.end_date.strftime("%d/%m/%Y")
        return f"({start_day}~{end_date})"

    def save_file(self, filename: str, dest_filepath: str = ".") -> None:
        self.writer.save_workbook(filepath=dest_filepath, filename=filename)
