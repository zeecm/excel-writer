from weekly_sales_stats.stats_creator import WeeklySalesStatsGenerator


def test_create_new_week_tab():
    sheet_location = (
        r"tests\test_files\Sale support statistics - 2023 (Wk 24)-beta.xlsx"
    )
    creator = WeeklySalesStatsGenerator(sheet_location)
    creator.create_new_week_tab()
    assert creator.writer.worksheets[0] == "Week 25"
