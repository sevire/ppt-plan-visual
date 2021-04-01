

visual_config = {
    'visual_config': {
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
    },
    'plotable_elements':
        [
            {
                'type': 'bar',
                'start_date': '27/3/2021',
                'end_date': '10/6/2021',
                'bar_height_in_tracks': 1,
                'format_category': 1,
            }
        ]
}


class PlanVisualDriver:
    def __init__(self, visual_driver_config):
        self.visual_driver_config = visual_driver_config

