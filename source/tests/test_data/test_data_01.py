from datetime import datetime

from pptx.util import Cm, Pt


def date_parse(text_date):
    return datetime.strptime(text_date, "%Y%m%d")


excel_plan_config_01 = {  # Jeremy's original PPT plan
    'excel_plan_sheet_name': 'ppt_plan_driver',
    'excel_plot_config_sheet_name': 'ppt_plot_config',
    'excel_format_config_sheet_name': 'ppt_format_config',
    'plan_start_row': 1
}

excel_plan_config_02 = {  # New from scratch slide for slide deck
    'excel_plan_sheet_name': 'UKViewPOAP',
    'excel_plot_config_sheet_name': 'ppt_plot_config',
    'excel_format_config_sheet_name': 'ppt_format_config',
    'plan_start_row': 1
}

excel_plan_config_03 = {  # New from scratch slide for slide deck
    'excel_plan_sheet_name': 'VisualPlanDriver',
    'excel_plot_config_sheet_name': 'ppt_plot_config',
    'excel_format_config_sheet_name': 'ppt_format_config',
    'plan_start_row': 2
}

excel_plan_config_smartsheet = {  # New from scratch slide for slide deck
    'excel_plan_sheet_name': 'UK-View Plan',
    'excel_plot_config_sheet_name': 'ppt_plot_config',
    'excel_format_config_sheet_name': 'ppt_format_config',
    'plan_start_row': 1
}

plot_area_config_ukview_01 = {  # Version of Jeremy's original PPT plan
    'top': Cm(3.21),
    'left': Cm(0.46),
    'bottom': Cm(10),
    'right': Cm(33.5),
    'track_height': Cm(0.5),
    'track_gap': Cm(0.2),
    'min_start_date': date_parse("20210101"),
    'max_end_date': date_parse("20220331"),
    'milestone_width': Cm(0.4),
    'milestone_text_width': Cm(5),
    'activity_text_width': Cm(5),  # Only used if text positioning is left or right (not shape)
}

plot_area_config_ukview_02 = {  # New from scratch slide for slide deck
    'top': Cm(3.86),
    'left': Cm(0),
    'bottom': Cm(0),
    'right': Cm(33.87),
    'track_height': Cm(0.5),
    'track_gap': Cm(0.2),
    'min_start_date': date_parse("20210101"),
    'max_end_date': date_parse("20220630"),
    'milestone_width': Cm(0.4),
    'milestone_text_width': Cm(5),
    'activity_text_width': Cm(5),  # Only used if text positioning is left or right (not shape)
}

plot_area_config_ukview_03 = {  # PPT view of smartsheets plan
    'top': Cm(3.21),
    'left': Cm(0.46),
    'bottom': Cm(10),
    'right': Cm(33.5),
    'track_height': Cm(0.5),
    'track_gap': Cm(0.2),
    'min_start_date': date_parse("20210101"),
    'max_end_date': date_parse("20220331"),
    'milestone_width': Cm(0.4),
    'milestone_text_width': Cm(5),
    'activity_text_width': Cm(5),  # Only used if text positioning is left or right (not shape)
}

format_config_01 = {
    'slide_level_categories': {
        'UKViewPOAP': {
            'swimlanes': [
                'Governance',
                'LI Data Team',
                'UK View Delivery',
                'APHA',
                'Devolved Authorities',
                'Technical Delivery'
            ],
            'swimlane_format_odd':
                {
                    'fill_rgb': (242, 242, 242),
                    'line_rgb': (242, 242, 242),
                    'font_size': Pt(14),
                    'font_bold': True,
                    'font_italic': False,
                    'font_colour_rgb': (166, 166, 166),
                    'text_align': 'left',
                    'text_vertical_align': 'top'
                },
            'swimlane_format_even':
                {
                    'fill_rgb': (255, 255, 255),
                    'line_rgb': (255, 255, 255),
                    'font_size': Pt(14),
                    'font_bold': True,
                    'font_italic': False,
                    'font_colour_rgb': (166, 166, 166),
                    'text_align': 'left',
                    'text_vertical_align': 'top'
}
        }
    },
    'format_categories': {
        'Governance Milestones 1': {
            'fill_rgb': (32, 56, 100),
            'line_rgb': (32, 56, 100),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',  # Values left, right or shape.  Shape only applies to bar and means align to shape
            'text_align': 'right',
            'text_vertical_align': 'middle'
},
        'Governance Milestones 1 (right)': {
            'fill_rgb': (32, 56, 100),
            'line_rgb': (32, 56, 100),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'right',  # Values left, right or shape.  Shape only applies to bar and means align to shape
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Governance Milestones 2': {
            'fill_rgb': (192, 0, 0),
            'line_rgb': (192, 0, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Governance Milestones 2 (right)': {
            'fill_rgb': (192, 0, 0),
            'line_rgb': (192, 0, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'right',
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Governance Milestones 3': {
            'fill_rgb': (255, 139, 12),
            'line_rgb': (255, 139, 12),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Governance Milestones 3 (right)': {
            'fill_rgb': (255, 139, 12),
            'line_rgb': (255, 139, 12),
            'corner_radius': Cm(0.1),
            'font_size': Pt(8),
            'font_bold': True,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'right',
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Activity Amber': {
            'fill_rgb': (255, 192, 0),
            'line_rgb': (255, 192, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'shape',
            'text_align': 'centre',
            'text_vertical_align': 'middle'
        },
        'Activity Amber (left)': {
            'fill_rgb': (255, 192, 0),
            'line_rgb': (255, 192, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Activity Amber (right)': {
            'fill_rgb': (255, 192, 0),
            'line_rgb': (255, 192, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'right',
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Activity Amber 2-Track': {
            'fill_rgb': (255, 192, 0),
            'line_rgb': (255, 192, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(16),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'centre',
            'text_vertical_align': 'middle'
},
        'Activity Red': {
            'fill_rgb': (255, 0, 0),
            'line_rgb': (255, 0, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'shape',  # Values left, right or shape.  Shape only applies to bar and means align to shape
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Activity Red (left)': {
            'fill_rgb': (255, 0, 0),
            'line_rgb': (255, 0, 0),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'shape',  # Values left, right or shape.  Shape only applies to bar and means align to shape
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Activity Green': {
            'fill_rgb': (146, 208, 80),
            'line_rgb': (146, 208, 80),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'shape',
            'text_align': 'shape',
            'text_vertical_align': 'middle'
        },
        'Activity Green (left)': {
            'fill_rgb': (146, 208, 80),
            'line_rgb': (146, 208, 80),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
        'Activity Green (right)': {
            'fill_rgb': (146, 208, 80),
            'line_rgb': (146, 208, 80),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'right',
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Activity Blue': {
            'fill_rgb': (0, 176, 240),
            'line_rgb': (0, 176, 240),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'shape',
            'text_align': 'left',
            'text_vertical_align': 'middle'
        },
        'Activity Blue (left)': {
            'fill_rgb': (0, 176, 240),
            'line_rgb': (0, 176, 240),
            'corner_radius': Cm(0.1),
            'font_size': Pt(9),
            'font_bold': False,
            'font_italic': False,
            'font_colour_rgb': (50, 50, 200),
            'text_position': 'left',
            'text_align': 'right',
            'text_vertical_align': 'middle'
        },
    }
}

plot_data = [
    {
        'id': 1,
        'description': 'This is activity 1',
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
        'description': 'This is activity 2',
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
        'description': 'This is activity 3',
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
        'description': 'This is activity 4',
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
        'description': 'This is activity 5',
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
        'description': 'This is activity 6',
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
        'description': 'This is activity 7',
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
        'description': 'This is activity 8',
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
        'description': 'This is activity 9',
        'type': 'bar',
        'start_date': date_parse('20210710'),
        'end_date': date_parse('20211003'),
        'swimlane': 'Swimlane 1',
        'track_num': 12,
        'bar_height_in_tracks': 1,
        'format_properties': "Amber",
    },
]
