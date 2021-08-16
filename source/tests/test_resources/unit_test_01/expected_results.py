"""
Expected results for full unit test.

This file sets out detailed expected results corresponding to a set of test files.

NOTE: Pathnames need to be relative because they need to run on the devops server as part of deployment of app.
"""
from source.tests.testing_utilities import Cm_to_ppt_points as cm2p

input_files_01 = {
        "visual_config": 'test_resources/unit_test_01/input_files/unit_test_01_config.xlsx',
        "excel_plan_file": 'test_resources/unit_test_01/input_files/unit_test_01.plan.xlsx',
        "plan_sheet_name": 'Plan',
        "ppt_template": "test_resources/unit_test_01/input_files/unit_test_dummy_ppt.pptx"  # We aren't creating a PPT file so shouldn't need this.
    }

width_per_month = cm2p(33.87)/31  # Just width of standard ppt template / days in Jan (could simplify!)

expected_results_01 = {
    "swimlanes": [
        # text, top, left, width, height
        ('SW1', 99, 99, 99, 99),
        ('SW2', 99, 99, 99, 99),
        ('SW3', 99, 99, 99, 99),
    ],
    "month_bars": [
        # text, top, left, width, height
        ('MB1', 99, 99, 99, 99),
        ('MB2', 99, 99, 99, 99),
        ('MB3', 99, 99, 99, 99),
    ],
    "plan_data": [
        # Expected results for each activity include formatting for two or three shapes, and text for one shape.
        # Type, text, [(top, left, width, height)] (x2 if done + to do)
        ("Activity", 'Activity 1', [
            (0, 0, 99, 99), # Graphic Shape 1 (no shape 2 for this case)
            (0, 0, 99, 99), # Text Shape
        ]),
        ("Activity", 'Activity 2', [
            (cm2p(0.1+0.5), 0, 99, 99),
            (cm2p(0.1+0.5), 0, 99, 99)
        ]),
        ("Activity", 'Activity 3', [
            (cm2p(0.1+0.5), round(width_per_month), 99, 99),
            (cm2p(0.1+0.5), round(width_per_month), 99, 99)
        ]),
    ]
}