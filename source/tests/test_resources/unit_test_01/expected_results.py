"""
Expected results for full unit test.

This file sets out detailed expected results corresponding to a set of test files.

NOTE: Pathnames need to be relative because they need to run on the devops server as part of deployment of app.
"""
from source.tests.testing_utilities import Cm_to_ppt_points as cm2p, parse_date

input_files_01 = {
        "visual_config": 'test_resources/unit_test_01/input_files/unit_test_01_config.xlsx',
        "excel_plan_file": 'test_resources/unit_test_01/input_files/unit_test_01_config.xlsx',
        "plan_sheet_name": 'Plan',
        "ppt_template": "test_resources/unit_test_01/input_files/unit_test_dummy_ppt.pptx"  # We aren't creating a PPT file so shouldn't need this.
    }

right = cm2p(33.87)
left = cm2p(2)  # Hard-coded but reflects what is in the input config file
top = cm2p(1)
width = right - left
width_per_day = width/31  # Just width of standard ppt template / days in Jan (could simplify!)

today = parse_date('2021-01-05')

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
        ("Activity", "Activity 01", [(360000, 720000, 11103097, 180000), (360000, 720000, 11103097, 180000)]),
        ("Activity", "Activity 02", [(576000, 720000, 11103097, 396000), (576000, 720000, 11103097, 396000)]),
        ("Activity", "Activity 03", [(1008000, 1090103, 10732994, 612000), (1008000, 1090103, 10732994, 612000)]),
        ("Activity", "Activity 04", [(1656000, 1460206, 740206, 180000), (1656000, 2200413, 5921652, 180000), (1656000, 1460206, 6661858, 180000)]),
    ]
}

