import logging
import sys
import time
from logging.handlers import RotatingFileHandler

from source.excel_config import ExcelFormatConfig, ExcelPlotConfig, ExcelSwimlaneConfig
from source.plan_visualiser import PlanVisualiser
from source.tests.test_data.test_data_01 import format_config_01


def main():
    root_logger = logging.getLogger()

    log_formatter = logging.Formatter("[%(levelname)-5.5s] %(asctime)s [%(threadName)-12.12s] %(message)s")

    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)

    # Probably doesn't need to rotate files as the log file is always created each time the app is run.
    file_handler = RotatingFileHandler(f"{'../logging'}/{'plan_to_ppt'}-{time_string}.log")
    file_handler.setFormatter(log_formatter)
    root_logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    root_logger.addHandler(console_handler)

    root_logger.setLevel(logging.INFO)

    root_logger.debug('Plan to PowerPoint plotting programme starting...')

    template_path = '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/UK-ViewPlanOnePager.pptx'

    # ts = time.gmtime()
    # time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)
    # logging.basicConfig(filename='../logging/create_plan_visual_{}.log'.format(time_string), level=logging.DEBUG)

    # If no command line arguments supplied, then assume running in test mode or
    # debug, and use hard-coded values for arguments

    # For testing, choose whether to use Excel import or test data
    source = "ExcelSmartSheeta"

    root_logger.info(f"Running from IDE, using fixed arguments, from {source}")

    plan_excel_config = {
        'excel_plan_sheet_name': 'UK-View Plan',
        'excel_plot_config_sheet_name': 'PlotConfig',
        'excel_format_config_sheet_name': 'FormatConfig',
        'plan_start_row': 1
    }
    plan_data_excel_file = '/Users/livestockinformation/Downloads/UK-View Plan.xlsx'
    root_logger.info(f'Using plan data from {plan_data_excel_file}')

    excel_config_path = '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx'
    root_logger.info(f'Using config info from {excel_config_path}')

    excel_plot_config = ExcelPlotConfig(excel_config_path, excel_sheet=plan_excel_config['excel_plot_config_sheet_name'])
    plot_area_config = excel_plot_config.parse_plot_config()

    excel_format_config = ExcelFormatConfig(excel_config_path, excel_sheet=plan_excel_config['excel_format_config_sheet_name'])
    format_config = {
        'format_categories': excel_format_config.parse_format_config()
    }

    swimlane_config = ExcelSwimlaneConfig(excel_config_path, excel_sheet='Swimlanes')
    swimlanes = swimlane_config.parse_swimlane_config()

    visualiser = PlanVisualiser.from_excelsmartsheet(plan_data_excel_file, plot_area_config, format_config,
                                                     plan_excel_config, template_path, swimlanes)

    visualiser.plot_slide()


if __name__ == '__main__':
    main()
