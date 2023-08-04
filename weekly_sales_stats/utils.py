from datetime import datetime, timedelta
from typing import NamedTuple


class DateRange(NamedTuple):
    start_date: datetime
    end_date: datetime


CURRENT_YEAR = 2023


def calculate_week_dates(week: int, year: int = CURRENT_YEAR) -> DateRange:
    first_monday_of_year = _calculate_first_monday_of_year(year)
    start_date_of_week = _calculate_start_date_of_week(first_monday_of_year, week)
    end_date_of_week = _calculate_end_date_of_week(start_date_of_week)
    return DateRange(start_date_of_week, end_date_of_week)


def _calculate_first_monday_of_year(year: int) -> datetime:
    reference_point = datetime(year, 1, 4)
    reference_point_weekday = reference_point.isocalendar().weekday
    days_to_first_monday = 1 - reference_point_weekday
    return reference_point + timedelta(days=days_to_first_monday)


def _calculate_start_date_of_week(
    first_monday_of_year: datetime, week: int
) -> datetime:
    return first_monday_of_year + timedelta(weeks=week - 1)


def _calculate_end_date_of_week(start_date_of_week: datetime) -> datetime:
    return start_date_of_week + timedelta(days=6)
