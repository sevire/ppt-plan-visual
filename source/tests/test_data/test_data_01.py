from datetime import datetime

from pptx.util import Cm, Pt


def date_parse(text_date):
    return datetime.strptime(text_date, "%Y%m%d")


# ToDo: Replace hard-coded swimlane order with configuration through Excel
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
        }
    }
}
