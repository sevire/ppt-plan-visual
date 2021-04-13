import logging
import sys
import time

from source.plan_visualiser import PlanVisualiser
from source.tests.test_data.test_data_01 import plot_data, excel_plan_config_01, \
    plot_area_config_ukview_01, format_config_01, excel_plan_config_02, plot_area_config_ukview_02, excel_plan_config_03

# Configuration for when using Excel to drive.  Until we implement command line arguments.
driver_data_set = {
    'ukview-jeremy_ppt': {
        'excel_plan_config ': excel_plan_config_01,
        'plot_area_config': plot_area_config_ukview_01,
        'format_config': format_config_01,
        'slide_level_config': format_config_01['slide_level_categories']['UKViewPOAP']
    },
    'ukview-poap': {
        'excel_plan_config': excel_plan_config_03,
        'plot_area_config': plot_area_config_ukview_02,
        'format_config': format_config_01,
        'slide_level_config': format_config_01['slide_level_categories']['UKViewPOAP']
    },
    'ukview-from-smartsheet': {
        'excel_plan_config ': excel_plan_config_01,
        'plot_area_config': plot_area_config_ukview_01,
        'format_config': format_config_01,
        'slide_level_config': format_config_01['slide_level_categories']['UKViewPOAP']
    }
}

data_set_to_use = driver_data_set['ukview-poap']
excel_path = '/Users/livestockinformation/Livestock Information Ltd/Data - Data Insights/UK View/planning/UKViewPOAP-01-Driver.xls'
template_path = '/Users/livestockinformation/Livestock Information Ltd/Data - Data Insights/UK View/planning/UK-ViewPlanOnePager.pptx'


def main():
    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)
    logging.basicConfig(filename='../logging/create_plan_visual_{}.log'.format(time_string), level=logging.DEBUG)

    # If no command line arguments supplied, then assume running in test mode or
    # debug, and use hard-coded values for arguments

    # For testing, choose whether to use Excel import or test data
    source = "Excel"

    if len(sys.argv) == 1:
        print(f"Running from IDE, using fixed arguments, from {source}")

        visualiser = None
        if source == "Excel":
            plan_data_excel_file = excel_path
            plan_excel_config = data_set_to_use['excel_plan_config']
            plot_area_config = data_set_to_use['plot_area_config']
            format_config = data_set_to_use['format_config']
            slide_level_config = data_set_to_use['slide_level_config']
            visualiser = PlanVisualiser.from_excel(plan_data_excel_file, plot_area_config, format_config, plan_excel_config, template_path, slide_level_config)
        elif source == "Test Data":
            plan_data = plot_data
            plot_config = driver_data_set['ukview-jeremy_ppt']['plot_area_config']
            visual_format_config = driver_data_set['ukview-jeremy_ppt']['format_config']
            slide_level_config = data_set_to_use['slide_level_config']
            visualiser = PlanVisualiser(plan_data, plot_config, visual_format_config, template_path, slide_level_config)
        else:
            print(f"Invalid source specified - {source}")

    else:
        plan_data_excel_file = sys.argv[1]
        plan_excel_config = sys.argv[2]
        ppt_template_path = sys.argv[3]
        visualiser = PlanVisualiser.from_excel(plan_data_excel_file, plan_excel_config, ppt_template_path)

    visualiser.plot_slide()


if __name__ == '__main__':
    main()
