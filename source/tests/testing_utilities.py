from datetime import datetime as dt

from dateutil.relativedelta import relativedelta


def parse_date(date_string):
    date = dt.strptime(date_string, '%Y-%m-%d')
    return date

def Cm_to_ppt_points(cm):
    return cm * 360000

def date_to_points(date, left, right, start_date, end_date):
    num_days = end_date - start_date + relativedelta(days=1)
    display_width = right - left

    scale = (date - start_date) / num_days
    width = scale * display_width

    rounded_width = round(width)

    return rounded_width
