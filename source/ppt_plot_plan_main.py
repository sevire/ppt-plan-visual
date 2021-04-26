import logging
import sys
import time

from source.excel_config import ExcelFormatConfig, ExcelPlotConfig
from source.plan_visualiser import PlanVisualiser
from source.tests.test_data.test_data_01 import format_config_01


def main():
    template_path = '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/UK-ViewPlanOnePager.pptx'

    # ts = time.gmtime()
    # time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)
    # logging.basicConfig(filename='../logging/create_plan_visual_{}.log'.format(time_string), level=logging.DEBUG)

    # If no command line arguments supplied, then assume running in test mode or
    # debug, and use hard-coded values for arguments

    # For testing, choose whether to use Excel import or test data
    source = "ExcelSmartSheeta"

    print(f"Running from IDE, using fixed arguments, from {source}")
    plan_excel_config = {
        'excel_plan_sheet_name': 'UK-View Plan',
        'excel_plot_config_sheet_name': 'PlotConfig',
        'excel_format_config_sheet_name': 'FormatConfig',
        'plan_start_row': 1
    }
    plan_data_excel_file = '/Users/livestockinformation/Downloads/UK-View Plan.xlsx'
    excel_config_path = '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx'

    excel_plot_config = ExcelPlotConfig(excel_config_path, excel_sheet=plan_excel_config['excel_plot_config_sheet_name'])
    plot_area_config = excel_plot_config.parse_plot_config()

    excel_format_config = ExcelFormatConfig(excel_config_path, excel_sheet=plan_excel_config['excel_format_config_sheet_name'])
    format_config = {
        'format_categories': excel_format_config.parse_format_config()
    }
    slide_level_config = format_config_01['slide_level_categories']['UKViewPOAP']
    visualiser = PlanVisualiser.from_excelsmartsheet(plan_data_excel_file, plot_area_config, format_config,
                                                     plan_excel_config, template_path, slide_level_config)

    visualiser.plot_slide()


if __name__ == '__main__':
    main()
