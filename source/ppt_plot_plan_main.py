import logging
import sys
import time

from source.plan_visual_driver import PlanVisualiser
from source.tests.test_data.test_data_01 import plot_data, format_config, template_path, excel_plan_config


def main():
    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)
    logging.basicConfig(filename='../logging/create_plan_visual_{}.log'.format(time_string), level=logging.DEBUG)

    # If no command line arguments supplied, then assume running in test mode or
    # debug, and use hard-coded values for arguments

    if len(sys.argv) == 1:
        print("Running from IDE, using fixed arguments")
        plan_data_excel_file = "/Users/livestockinformation/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/excel_plan_file/ExcelPlanFile01.xls"
        plan_excel_config = excel_plan_config
        ppt_template_path = template_path
    else:
        plan_data_excel_file = sys.argv[1]
        plan_excel_config = sys.argv[2]
        ppt_template_path = sys.argv[3]

    visualiser = PlanVisualiser.from_excel(plan_data_excel_file, plan_excel_config, ppt_template_path)
    visualiser.plot_slide()


if __name__ == '__main__':
    main()
