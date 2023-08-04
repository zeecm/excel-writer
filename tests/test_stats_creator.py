import os

from weekly_sales_stats.stats_creator import WeeklySalesStatsGenerator


def test_create_new_week_tab():
    sheet_location = os.path.join("tests", "test_files", "sample_weekly_stats.xlsx")
    creator = WeeklySalesStatsGenerator(sheet_location)
    creator.create_new_week_tab()
    assert creator.writer.worksheets[0] == "Week 25"
