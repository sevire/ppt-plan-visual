from datetime import datetime

from dateutil.utils import today

from source.visualiser.visual_element_shape import VisualElementShape


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
        if min_start_date is None:
            self.min_start_date = None  # Will get calculated from plan data later
        else:
            self.min_start_date = min_start_date

        max_end_date = plot_config['max_end_date']
        if max_end_date is None:
            self.max_end_date = None  # Will get calculated from plan data later
        else:
            self.max_end_date = max_end_date

        self.plot_area_width = self.right - self.left

    def date_to_x_coordinate(self, date, alignment_case="start"):
        """
        Calculates the x coordinate within a PowerPoint slide of a specific date.

        Note we may want a slightly different outcome depending upon the scenario, driven by the fact that a
        day takes up a finite amount of space and sometimes we want the left edge and sometimes we want the
        right edge, and sometimes we want the middle.

        Specifically:

        - If we are plotting the left hand edge of an activity bar then the point should be at the start of the day
        - If we are plotting the right hand edge of an activity bar then the point should be at the end of the day
        - If we are plotting a date line then the line should (probably) be in the middle of the day.

        The calculation works as follows:

        Case = "start"
        - Work out how many whole days there are before the day to be plotted.
        - The result is 1 more than this (in the units of PPT).

        Case = "end"
        - Work out one more than the number of whole days before the day to be plotted.
        - The result is 1 less than this (in the units of PPT)

        Case = "middle"
        - The result is 0.5 days more than the number of whole days before the day to be plotted.

        :param date:
        :param alignment_case: Values are "start" (default), "end", "middle"
        :return:
        """
        if alignment_case == "start":
            num_days = date.toordinal() - self.min_start_date.toordinal()
            additional_units = 0  # 1 unit, not one Cm (1/360000 of Cm)
        elif alignment_case == "end":
            num_days = date.toordinal() - self.min_start_date.toordinal() + 1
            additional_units = 0
        else:  # assume case is "middle"
            num_days = date.toordinal() - self.min_start_date.toordinal()
            additional_units = self.width_of_one_day()/2

        main = (num_days / self.num_days_in_date_range) * self.plot_area_width
        distance_from_left_of_plot_area = main + additional_units

        x_coord = self.left + distance_from_left_of_plot_area

        return x_coord

    def track_number_to_y_coordinate(self, track_num):
        return self.top + (track_num - 1) * (self.track_height + self.track_gap)

    def height_of_track(self, num_tracks):
        return num_tracks * self.track_height + (num_tracks-1) * self.track_gap

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

        left = self.date_to_x_coordinate(start_date, "start")
        right = self.date_to_x_coordinate(end_date, "end") - gap_increment
        width = right - left

        return left, right, width

    def milestone_left(self, start_date, milestone_width):
        left = int(self.date_to_x_coordinate(start_date, "middle") - (milestone_width / 2))

        return left
