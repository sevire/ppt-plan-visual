from datetime import datetime

import pandas as pd
from dateutil.utils import today

from source.refactor_temp.visual_element_shape import VisualElementShape
from source.utilities import day_increment


class PlotDriver:
    """
    Used to translate plan type data into data which can be used to plot shapes.
    Example - convert from a date to a horizontal position on the slide.

    """
    def __init__(self, plot_config):
        self.top = plot_config['top']
        self.left = plot_config['left']
        self.bottom = plot_config['bottom']
        self.right = plot_config['right']
        self.track_height = plot_config['track_height']
        self.track_gap = plot_config['track_gap']
        self.min_activity_text_width = plot_config['activity_text_width']
        self.milestone_text_width = plot_config['milestone_text_width']
        self.text_margin = plot_config['text_margin']
        self.milestone_width = plot_config['milestone_width']
        self.activity_shape = VisualElementShape[plot_config['activity_shape'].upper()]
        self.milestone_shape = VisualElementShape[plot_config['milestone_shape'].upper()]

        if 'today' in plot_config:
            self.today = plot_config['today']
        else:
            self.today = today()

        min_start_date = plot_config['min_start_date']
        if pd.isnull(min_start_date):
            self.min_start_date = None  # Will get calculated from plan data later
        else:
            self.min_start_date = min_start_date

        max_end_date = plot_config['max_end_date']
        if pd.isnull(max_end_date):
            self.max_end_date = None  # Will get calculated from plan data later
        else:
            self.max_end_date = max_end_date

        self.plot_area_width = self.right - self.left

    def date_to_x_coordinate(self, date):
        """
        This isn't quite as easy as it seems if we do it to pinpoint accuracy as the fact that
        different months have different numbers of days will come into play.

        For now just calculate based on how many days through the range the date is and
        calculate distance as a proportion of total plot width.

        In fact this approach would be completely accurate if we were using days as the unit
        not months.

        :param date:
        :return:
        """
        num_days_from_min_date = date.toordinal() - self.min_start_date.toordinal() + 1
        proportion_of_plot_width = num_days_from_min_date / self.num_days_in_date_range
        distance_from_left_of_plot_area = proportion_of_plot_width * self.plot_area_width
        x_coord = round(self.left + distance_from_left_of_plot_area)

        return x_coord

    def track_number_to_y_coordinate(self, track_num):
        return round(self.top + (track_num - 1) * (self.track_height + self.track_gap))

    def height_of_track(self, num_tracks):
        return round(num_tracks * self.track_height + (num_tracks-1) * self.track_gap)

    @staticmethod
    def string_date_to_date(string_date):
        return datetime.strptime(string_date, "%Y%m%d").date()

    def width_of_one_day(self):
        proportion_of_plot_width = 1 / self.num_days_in_date_range
        distance_from_left_of_plot_area = proportion_of_plot_width * self.plot_area_width

        return distance_from_left_of_plot_area

    def shape_parameters(self, start_date, end_date, gap_flag=True):
        """
        Calculates the (x-axis) parameters required for plotting a shape based on a start and end date.
        To allow a visible but very small gap between activities, the right hand edge is brought in by a fixed number of
        units, but to allow for cases where this isn't desirable (e.g. when plotting month bars or the done and to do
        parts of an activity, when gap_flag is False this adjustment won't be made.

        :param start_date:
        :param end_date:
        :param gap_flag:
        :return:
        """
        if gap_flag is True:
            gap_increment = 15000
        else:
            gap_increment = 0

        left = self.date_to_x_coordinate(start_date)
        right = self.date_to_x_coordinate(day_increment(end_date, 1)) - gap_increment
        width = right - left

        return left, right, width

    def milestone_left(self, start_date, milestone_width):
        left = int(self.date_to_x_coordinate(start_date) + (self.width_of_one_day() / 2) - (milestone_width / 2))

        return left
