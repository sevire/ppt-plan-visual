from datetime import datetime

from pptx.util import Cm


def date_parse(text_date):
    return datetime.strptime(text_date, "%Y%m%d")


excel_plan_config = {
    'excel_plan_sheet_name': 'ppt_plan_driver',
    'excel_plot_config_sheet_name': 'ppt_plot_config',
    'excel_format_config_sheet_name': 'ppt_format_config',
    'plan_start_row': 1
}

plot_area_config = {
    'top': Cm(4.56),
    'left': Cm(2.33),
    'bottom': Cm(10),
    'right': Cm(29.21+2.33),
    'track_height': Cm(0.5),
    'track_gap': Cm(0.2),
    'min_start_date': date_parse("20210301"),
    'max_end_date': date_parse("20220623"),
    'milestone_width': Cm(0.5)
}

format_config = {
    'format_categories': {
        'Red': {
            'fill_rgb': (199, 66, 33),
            'line_rgb': (181, 60, 31),
            'corner_radius': Cm(0.1)
        },
        'Amber': {
            'fill_rgb': (255, 101, 34),
            'line_rgb': (207, 81, 29),
            'corner_radius': Cm(0.2)
        },
    }
}

plot_data = [
    {
        'id': 1,
        'type': 'bar',
        'start_date': date_parse('20210301'),
        'end_date': date_parse('20220623'),
        'swimlane': 'Swimlane 1',
        'track_num': 1,
        'bar_height_in_tracks': 2,
        'format_properties': "Red",
    },
    {
        'id': 2,
        'type': 'milestone',
        'start_date': date_parse('20210410'),
        'end_date': date_parse('20210703'),
        'swimlane': 'Swimlane 1',
        'track_num': 3,
        'bar_height_in_tracks': 1,
        'format_properties': "Red",
    },
    {
        'id': 3,
        'type': 'bar',
        'start_date': date_parse('20210510'),
        'end_date': date_parse('20210803'),
        'swimlane': 'Swimlane 1',
        'track_num': 4,
        'bar_height_in_tracks': 1,
        'format_properties': "Amber",
    },
    {
        'id': 4,
        'type': 'bar',
        'start_date': date_parse('20210610'),
        'end_date': date_parse('20210903'),
        'swimlane': 'Swimlane 1',
        'track_num': 5,
        'bar_height_in_tracks': 3,
        'format_properties': "Amber",
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210910'),
        'end_date': date_parse('20211203'),
        'swimlane': 'Swimlane 1',
        'track_num': 5,
        'bar_height_in_tracks': 1,
        'format_properties': "Amber",
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210910'),
        'end_date': date_parse('20211203'),
        'swimlane': 'Swimlane 1',
        'track_num': 6,
        'bar_height_in_tracks': 1,
        'format_properties': "Red",
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210910'),
        'end_date': date_parse('20211203'),
        'swimlane': 'Swimlane 1',
        'track_num': 7,
        'bar_height_in_tracks': 1,
        'format_properties': "Red",
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 11,
        'bar_height_in_tracks': 1,
        'format_properties': "Amber",
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 12,
        'bar_height_in_tracks': 1,
        'format_properties': "Amber",
    },
]

template_path = '/Users/livestockinformation/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/' \
                'ppt_templates/PlanVisual-01.pptx'
