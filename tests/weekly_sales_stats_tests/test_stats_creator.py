import os

import pytest

from weekly_sales_stats.stats_creator import WeeklySalesStatsGenerator


class TestWeeklySalesStatsGenerator:
    def setup_method(self):
        sheet_location = os.path.join("tests", "test_files", "sample_weekly_stats.xlsx")
        self.generator = WeeklySalesStatsGenerator(sheet_location)

    def test_create_new_week_tab(self):
        self.generator.create_new_week_tab()
        assert self.generator.writer.worksheet_names[0] == "Week 25"

    def test_set_data_summary_date(self):
        self.generator.create_new_week_tab()
        self.generator._set_data_summary_date()
        cell = self.generator.writer.cell(0, "D2")
        assert cell.value == "Week 25 (19~25/06/2023)"

    def test_new_tab_generated_from_template(self):
        self.generator.create_new_week_tab()
        assert self.generator.writer.cell(0, "D2").value == "<week_date>"
