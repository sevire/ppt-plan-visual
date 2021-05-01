import logging
import time
from logging.handlers import RotatingFileHandler
from source.excel_config import ExcelFormatConfig, ExcelPlotConfig, ExcelSwimlaneConfig
from source.excel_plan import ExcelPlan
from source.plan_visualiser import PlanVisualiser


def get_parameters():
    parameters = {
        'excel_plan_file': '/Users/livestockinformation/Downloads/UK-View Plan.xlsx',
        'excel_plan_sheet': 'UK-View Plan',
        'excel_plot_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx',
        'excel_plot_cfg_sheet': 'PlotConfig',
        'excel_format_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx',
        'excel_format_cfg_sheet': 'FormatConfig',
        'swimlanes_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx',
        'swimlanes_cfg_sheet': 'Swimlanes',
        'ppt_template_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK Data/UK View/planning/planning-visual/UK-ViewPlanOnePager.pptx',
    }

    return parameters


def configure_logger(logger):
    log_formatter = logging.Formatter("[%(levelname)-5.5s] %(asctime)s [%(threadName)-12.12s] %(message)s")

    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)

    # Probably doesn't need to rotate files as the log file is always created each time the app is run.
    file_handler = RotatingFileHandler(f"{'../logging'}/{'plan_to_ppt'}-{time_string}.log")
    file_handler.setFormatter(log_formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    logger.setLevel(logging.INFO)


def main():
    root_logger = logging.getLogger()
    configure_logger(root_logger)

    root_logger.debug('Plan to PowerPoint plotting programme starting...')
    root_logger.info(f"Running from IDE, using fixed arguments")

    parameters = get_parameters()

    excel_plan_file = parameters['excel_plan_file']
    excel_plan_sheet = parameters['excel_plan_sheet']
    excel_plot_cfg_file = parameters['excel_plot_cfg_file']
    excel_plot_cfg_sheet = parameters['excel_plot_cfg_sheet']
    excel_format_cfg_file = parameters['excel_format_cfg_file']
    excel_format_cfg_sheet = parameters['excel_format_cfg_sheet']
    swimlanes_cfg_file = parameters['swimlanes_cfg_file']
    swimlanes_cfg_sheet = parameters['swimlanes_cfg_sheet']
    ppt_template_file = parameters['ppt_template_file']

    root_logger.info(f'Using plan data from {excel_plan_file}')
    extracted_plan_data = ExcelPlan.read_plan_data(excel_plan_file, excel_plan_sheet)

    plot_config_object = ExcelPlotConfig(excel_plot_cfg_file, excel_sheet=excel_plot_cfg_sheet)
    plot_area_config = plot_config_object.parse_plot_config()

    excel_format_config_object = ExcelFormatConfig(excel_format_cfg_file, excel_sheet=excel_format_cfg_sheet)
    format_config = excel_format_config_object.parse_format_config()

    swimlane_config_object = ExcelSwimlaneConfig(swimlanes_cfg_file, excel_sheet=swimlanes_cfg_sheet)
    swimlanes = swimlane_config_object.parse_swimlane_config()

    visualiser = PlanVisualiser(
        extracted_plan_data,
        plot_area_config,
        format_config,
        ppt_template_file,
        swimlanes
    )

    visualiser.plot_slide()


if __name__ == '__main__':
    main()
