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
    'max_end_date': date_parse("20220623")
}

format_config = {
    'format_categories': {
        'category-01': {
            'bar_shape': 'rectangle',
            'bar_fill_rgb': (3,3,3),
            'bar_line_rgb': (4,4,4),
            'bar_line_thickness': '???'
        },
        'category-02': {
            'bar_shape': 'rounded_rectangle',
            'corner_radius': '???',
            'bar_fill_rgb': (3, 3, 3),
            'bar_line_rgb': (4, 4, 4),
            'bar_line_thickness': '???'
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
        'format_properties': 1,
    },
    {
        'id': 2,
        'type': 'bar',
        'start_date': date_parse('20210410'),
        'end_date': date_parse('20210703'),
        'swimlane': 'Swimlane 1',
        'track_num': 3,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 3,
        'type': 'bar',
        'start_date': date_parse('20210510'),
        'end_date': date_parse('20210803'),
        'swimlane': 'Swimlane 1',
        'track_num': 4,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 4,
        'type': 'bar',
        'start_date': date_parse('20210610'),
        'end_date': date_parse('20210903'),
        'swimlane': 'Swimlane 1',
        'track_num': 5,
        'bar_height_in_tracks': 3,
        'format_properties': 1,
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 8,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 9,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 10,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 11,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
    {
        'id': 5,
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 12,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
]

template_path = '/Users/livestockinformation/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/' \
                'ppt_templates/PlanVisual-01.pptx'
