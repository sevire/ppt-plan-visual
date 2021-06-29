from unittest import TestCase

from colour import Color
from pptx.util import Cm

from source.plot_driver import PlotDriver
from source.refactor_temp.activity_layout_attributes import ActivityLayoutAttributes
from source.refactor_temp.plan_activity import PlanActivity
from source.refactor_temp.shape_formatting import ShapeFormatting
from source.refactor_temp.text_formatting import TextFormatting
from source.refactor_temp.visual_element_shape import VisualElementShape
from source.tests.testing_utilities import parse_date

test_case_01 = {
    'visual_start_date': parse_date('2021-01-01'),
    'visual_end_date': parse_date('2021-07-31'),
    'visual_left': Cm(0),
    'visual_right': Cm(33.87),
    'activity_start_date': parse_date('2021-01-15'),
    'activity_end_date': parse_date('2021-07-27'),
    'expected_left': 999,
    'expected_width': 999,
    'today': parse_date('2021-06-29'),
    'num_days_in_date_range': 212
}


class DummyShapes:
    def add_shape(self, shape, top, left, width, height):
        self.left = left,
        self.width = width


class TestHorizontalLayout(TestCase):
    def test_horizontal_layout(self):
        layout_attributes = ActivityLayoutAttributes(
            'Swimlane-01',
            1,
            1,
            "shape"
        )
        display_shape = VisualElementShape.RECTANGLE
        plan_visual_config = PlotDriver(
            {
                'top': Cm(3.86),
                'left': test_case_01['visual_left'],
                'bottom': Cm(20),
                'right': Cm(30),
                'track_height': Cm(0.6),
                'track_gap': Cm(0.2),
                'activity_text_width': Cm(5),
                'milestone_text_width': Cm(5),
                'text_margin': Cm(0.1),
                'milestone_width': Cm(0.4),
                'activity_shape': 'rectangle',
                'milestone_shape': 'diamond',
                'min_start_date': parse_date('2021-01-01'),
                'max_end_date': parse_date('2021-07-31')
            }
        )
        plan_visual_config.num_days_in_date_range = test_case_01['num_days_in_date_range']
        text_formatting = TextFormatting()
        formatting_1 = ShapeFormatting(
            Color(rgb=(1,1,1)),
            Color(rgb=(1,1,1)),
            None,
            text_formatting,
        )

        activity = PlanActivity(
            12345,
            'Dummy',
            'bar',
            test_case_01['activity_start_date'],
            test_case_01['activity_end_date'],
            layout_attributes,
            display_shape,
            plan_visual_config,
            formatting_1,
            formatting_1,
            test_case_01['today'],
            1
        )

        dummy_shapes = DummyShapes()
        activity.plot_ppt_shapes(dummy_shapes)

        self.assertEqual(test_case_01['expected_left'], dummy_shapes.left)