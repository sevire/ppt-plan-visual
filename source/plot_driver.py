from datetime import datetime

import pandas as pd


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

        try:
            num_days_from_min_date = date.toordinal() - self.min_start_date.toordinal()
        except:
            pass
        proportion_of_plot_width = num_days_from_min_date / self.num_days_in_date_range
        distance_from_left_of_plot_area = proportion_of_plot_width * self.plot_area_width
        x_coord = self.left + distance_from_left_of_plot_area

        return x_coord

    def track_number_to_y_coordinate(self, track_num):
        return self.top + (track_num - 1) * (self.track_height + self.track_gap)

    def height_of_track(self, num_tracks):
        return num_tracks * self.track_height + (num_tracks-1) * self.track_gap

    def string_date_to_date(self, string_date):
        return datetime.strptime(string_date, "%Y%m%d").date()

    def width_of_one_day(self):
        proportion_of_plot_width = 1 / self.num_days_in_date_range
        distance_from_left_of_plot_area = proportion_of_plot_width * self.plot_area_width

        return distance_from_left_of_plot_area
