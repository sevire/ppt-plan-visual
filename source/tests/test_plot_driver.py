from unittest import TestCase
from ddt import ddt, data, unpack
from pptx.util import Cm

from source.visualiser.plot_driver import PlotDriver
from source.tests.testing_utilities import parse_date, date_to_points

tpc = {  # Test Plot Config
    'top': Cm(0),
    'left': Cm(0),
    'bottom': Cm(10),
    'right': Cm(20),
    'track_height': Cm(1),
    'track_gap': Cm(0.5),
    'min_start_date': parse_date('2021-01-01'),
    'max_end_date': parse_date('2021-01-31'),
    'activity_text_width': Cm(5),
    'milestone_text_width': Cm(5),
    'text_margin': Cm(0.1),
    'activity_width': Cm(5),
    'milestone_width': Cm(5),
    'activity_shape': 'RECTANGLE',
    'milestone_shape': 'DIAMOND'
}


def points(string_date):
    date = parse_date(string_date)
    return date_to_points(date, left=tpc['left'], right=tpc['right'], start_date=tpc['min_start_date'], end_date=tpc['max_end_date'])


one_cm = 360_000  # This is how PPT measures distances.

date_to_x_test_data = [
    ('2021-01-01', 0),
    ('2021-01-31', 6_967_742)
]

plot_activity_test_data = [
    ('2021-01-01', '2021-01-31', 0, 7_200_000)
]

plot_milestone_test_data = [
    ('2021-01-01', Cm(0.5), 26_129)
]


def plot_activity_test_generator():
    for entry in plot_activity_test_data:
        string_start, string_end, exp_left, exp_width = entry
        yield "left", parse_date(string_start), parse_date(string_end), exp_left
        yield "width", parse_date(string_start), parse_date(string_end), exp_width


def plot_milestone_test_generator():
    for entry in plot_milestone_test_data:
        string_start, milestone_width, exp_left = entry
        yield parse_date(string_start), milestone_width, exp_left


@ddt
class TestPlotDriver(TestCase):
    def setUp(self) -> None:
        self.plot_driver = PlotDriver(tpc)
        self.plot_driver.num_days_in_date_range = self.plot_driver.max_end_date.toordinal() - self.plot_driver.min_start_date.toordinal() + 1

    @data(*date_to_x_test_data)
    @unpack
    def test_date_to_x_coordinate(self, string_test_date, expected_x):
        test_date = parse_date(string_test_date)
        x = self.plot_driver.date_to_x_coordinate(test_date)
        self.assertEqual(expected_x, x)

    @data(*plot_activity_test_generator())
    @unpack
    def test_plot_activity(self, test_type, start_date, end_date, expected):

        left, _, width = self.plot_driver.shape_parameters(start_date, end_date, gap_flag=False)

        if test_type == "left":
            self.assertEqual(expected, left)
        elif test_type == "width":
            self.assertEqual(expected, width)
        else:
            self.fail(f"Unexpected test type {test_type}")

    @data(*plot_milestone_test_generator())
    @unpack
    def test_plot_milestone(self, start_date, milestone_width, expected_left):

        left = self.plot_driver.milestone_left(start_date, milestone_width)

        self.assertEqual(expected_left, left)
