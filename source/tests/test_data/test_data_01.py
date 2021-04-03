from datetime import datetime

from pptx.util import Cm

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
    'min_start_date': "20210301",
    'max_end_date': "20220623"
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
        'start_date': '20210310',
        'end_date': '20210610',
        'track_num': 5,
        'bar_height_in_tracks': 2,
        'format_properties': 1,
    },
    {
        'id': 1,
        'type': 'bar',
        'start_date': '20210410',
        'end_date': '20210703',
        'track_num': 7,
        'bar_height_in_tracks': 1,
        'format_properties': 1,
    },
]

template_path = '/Users/livestockinformation/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/' \
                'ppt_templates/PlanVisual-01.pptx'
