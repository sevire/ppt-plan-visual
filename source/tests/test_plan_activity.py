from unittest import TestCase

from colour import Color
from ddt import ddt, data, unpack
from pptx.util import Cm

from source.refactor_temp.shape_formatting import ShapeFormatting
from source.refactor_temp.activity_layout_attributes import ActivityLayoutAttributes
from source.refactor_temp.plan_activity import PlanActivity
from source.plot_driver import PlotDriver
from source.tests.testing_utilities import parse_date
from source.refactor_temp.visual_element_shape import VisualElementShape

plan_visual_config_test_data = {
    'vis_cfg_01': {
        'top': Cm(0),
        'left': Cm(0),
        'bottom': Cm(20),
        'right': Cm(30),
        'track_height': Cm(1),
        'track_gap': Cm(0.5),
        'min_start_date': None,
        'max_end_date': None,
        'milestone_width': None,
        'milestone_text_width': Cm(0.5),
        'activity_text_width': Cm(5),
        'text_margin': Cm(0.2),
    }
}
display_attribute_test_data = {
    'disp_01': {
        'line_colour': Color(rgb=(0,0,0)),
        'fill_colour': Color(rgb=(0,0,0)),
        'font_colour': Color(rgb=(0,0,0)),
        'text_layout': Color(rgb=(0,0,0))
    }
}

layout_attributes_test_records = {
    'layout_01': {
        'swimlane_name': 'yyy',
        'track_number': 'yyy',
        'number_of_tracks_to_span': 'yyy',
        'text_layout': 'yyy'
    }
}

swimlane_test_data_records = {
    'swim_01': {
        'Swimlane 1': {
            'start_track': 1,
            'end_track': 5
        }
    }
}

activity_test_data = {
    'act_01': {
        'activity_id': 12345,
        'description': 'Test activity 01',
        'activity_type': 'activity',
        'start_date': parse_date('2021-01-01'),
        'end_date': parse_date('2021-01-31'),
        'display_shape': VisualElementShape.ROUNDED_RECTANGLE,
    }
}


test_records = [
    (
        activity_test_data['act_01'],
        layout_attributes_test_records['layout_01'],
        display_attribute_test_data['disp_01'],
        display_attribute_test_data['disp_01'],
        plan_visual_config_test_data['vis_cfg_01'],
        swimlane_test_data_records['swim_01'],
        {  # Expected results record for this combination of test data
            'is_past': True,
            'is_future': False,
            'is_current': False,
            'plot_start': 0,
        }
    )
]


@ddt
class TestPlanActivity(TestCase):
    @data(*test_records)
    @unpack
    def test_create_plan_activity(
            self,
            activity_data,
            layout_data,
            display_data,
            done_display_data,
            vis_cfg,
            swimlane_data,
            exp_res
    ):
        layout_properties = ActivityLayoutAttributes(
            layout_data['swimlane_name'],
            layout_data['track_number'],
            layout_data['number_of_tracks_to_span'],
            layout_data['text_layout']
        )

        display_properties = ShapeFormatting(
            line_colour=display_data['line_colour'],
            fill_colour=display_data['line_colour'],
            font_colour=display_data['font_colour'],
            text_layout=display_data['text_layout']
        )

        done_display_properties = ShapeFormatting(
            line_colour=done_display_data['line_colour'],
            fill_colour=done_display_data['line_colour'],
            font_colour=done_display_data['font_colour'],
            text_layout=done_display_data['text_layout']
        )
        plot_driver = PlotDriver(vis_cfg)

        activity = PlanActivity(
            activity_data['activity_id'],
            activity_data['description'],
            activity_data['activity_type'],
            activity_data['start_date'],
            activity_data['end_date'],
            layout_properties,
            activity_data['display_shape'],
            plot_driver,
            swimlane_data,
            display_properties,
            done_display_properties,
        )

        self.assertEqual(exp_res['is_past'], activity.is_past())
        self.assertEqual(exp_res['is_future'], activity.is_future())
        self.assertEqual(exp_res['is_current'], activity.is_current())
        self.assertTrue(exp_res['plot_start'], activity._plot_left)
