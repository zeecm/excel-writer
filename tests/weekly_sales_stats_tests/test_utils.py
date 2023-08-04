from datetime import datetime

import pytest

from weekly_sales_stats.utils import DateRange, calculate_week_dates


@pytest.mark.parametrize(
    "year, week, expected_range",
    [
        (2023, 24, DateRange(datetime(2023, 6, 12), datetime(2023, 6, 18))),
        (2023, 25, DateRange(datetime(2023, 6, 19), datetime(2023, 6, 25))),
    ],
)
def test_calculate_week_dates(year: int, week: int, expected_range: DateRange):
    daterange = calculate_week_dates(week, year)
    assert daterange == expected_range
