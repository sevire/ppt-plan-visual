import os
from calendar import monthrange
from datetime import date

from dateutil.relativedelta import relativedelta


def get_path_name_ext(path):
    folder = os.path.dirname(path)
    file = os.path.basename(path)
    base, ext = os.path.splitext(file)

    return folder, base, ext


def first_day_of_month(date):
    return date.replace(day=1)


def last_day_of_month(date):
    return date.replace(day=monthrange(date.year, date.month)[1])


def month_increment(date, num_months):
    return date + relativedelta(months=num_months)


def day_increment(in_date, num_days):
    return in_date + relativedelta(days=num_days)


def num_months_between_dates(start_date, end_date):
    r = relativedelta(end_date, start_date)
    return r.years * 12 + r.months + 1


def iterate_months(date, num_months):
    yield date
    for month in range(num_months):
        yield month_increment(date, month)


def is_current(start_date, end_date):
    today = date.today()
    return start_date < today < end_date


def is_past(start_date, end_date):
    today = date.today()
    return start_date < end_date <= today


def is_future(start_date, end_date):
    today = date.today()
    return today < start_date < end_date


def is_nan(value: str):
    # Uses property that nan != nan in pandas/numpy
    ret_value = (value != value)
    return ret_value


class SwimlaneManager:
    """
    - Manages list of swimlanes which drives positioning of swimlanes on visual.
    - Takes list of swimlanes and uses order in list to determine swimlane number.
    - Implements method to return swimlane number to user during visual creation.
      If a request for the number of a non-existent swimlane is made, the class will add the swimlane to the end of the list
      and return it's implied number, to ensure consistency.
    """

    def __init__(self, swimlane_data):
        self.swimlane_data = swimlane_data

    def get_swimlane_number(self, swimlane_name):
        # Make sure there is an entry for this swimlane.  If not already there, add it.
        if swimlane_name not in self.swimlane_data:
            self.swimlane_data.append(swimlane_name)

        # Filter out all entries which aren't the indicated swimlane.
        # If there is a duplicate, ignore and just return the lowest.  That will be consistent.
        this_swimlane_only = [index for index, swimlane in enumerate(self.swimlane_data) if swimlane == swimlane_name]
        return this_swimlane_only[0] + 1
