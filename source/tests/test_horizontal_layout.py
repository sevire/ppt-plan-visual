from unittest import TestCase
from colour import Color
from pptx import Presentation
from pptx.util import Cm
from source.plot_driver import PlotDriver
from source.refactor_temp.activity_layout_attributes import ActivityLayoutAttributes
from source.refactor_temp.plan_activity import PlanActivity
from source.refactor_temp.shape_formatting import ShapeFormatting
from source.refactor_temp.text_formatting import TextFormatting
from source.refactor_temp.visual_element_shape import VisualElementShape
from source.tests.testing_utilities import parse_date
from unittest.mock import Mock

test_case_01 = {
    'parameters': {
        'visual_start_date': parse_date('2021-01-01'),
        'visual_end_date': parse_date('2021-07-31'),
        'visual_left': Cm(0),
        'visual_right': Cm(33.87),
        'activity_start_date': parse_date('2021-01-15'),
        'activity_end_date': parse_date('2021-07-27'),
        'today': parse_date('2021-06-29'),
        'num_days_in_date_range': 212
    },
    'expected_results': {  # Calculations are from Excel
        'expected_num_shapes': 2,
        'expected_left_1': 805212,
        'expected_width_1': 9489991,
        'expected_left_2': 10295203,
        'expected_width_2': 1667938,
    }
}


class TestHorizontalLayout(TestCase):
    def test_horizontal_layout(self):
        layout_attributes = ActivityLayoutAttributes(
            'Swimlane-01',
            1,
            1,
            "shape"
        )
        display_shape = VisualElementShape.RECTANGLE
        parameters = test_case_01['parameters']
        expected_results = test_case_01['expected_results']
        plan_visual_config = PlotDriver(
            {
                'top': Cm(3.86),
                'left': parameters['visual_left'],
                'bottom': Cm(20),
                'right': Cm(33.87),
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
        plan_visual_config.num_days_in_date_range = parameters['num_days_in_date_range']
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
            parameters['activity_start_date'],
            parameters['activity_end_date'],
            layout_attributes,
            display_shape,
            plan_visual_config,
            formatting_1,
            formatting_1,
            parameters['today'],
            1
        )

        pres = Presentation()
        slide_layout = pres.slide_layouts[0]
        slide = pres.slides.add_slide(slide_layout)
        shapes = slide.shapes

        shapes = activity.plot_ppt_shapes(shapes)
        actual_left_1 = shapes[0].left
        actual_width_1 = shapes[0].width
        actual_left_2 = shapes[1].left
        actual_width_2 = shapes[1].width
        self.assertEqual(expected_results['expected_num_shapes'], len(shapes))
        self.assertEqual(expected_results['expected_left_1'], actual_left_1)
        self.assertEqual(expected_results['expected_width_1'], actual_width_1)
        self.assertEqual(expected_results['expected_left_2'], actual_left_2)
        self.assertEqual(expected_results['expected_width_2'], actual_width_2)