import logging
import sys
import time
from logging.handlers import RotatingFileHandler
from source.excel_config import ExcelFormatConfig, ExcelPlotConfig, ExcelSwimlaneConfig
from source.excel_plan import ExcelPlan
from source.plan_visualiser import PlanVisualiser

root_logger = logging.getLogger()

# No args provided so use hard-coded defaults
parameters_01 = {
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
parameters_02 = {
    'excel_plan_file': '/Users/livestockinformation/Downloads/LI-DeliveryScenarios.xlsx',
    'excel_plan_sheet': 'LI-DeliveryScenarios',
    'excel_plot_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/UKV-ScenariosConfig.xlsx',
    'excel_plot_cfg_sheet': 'PlotConfig',
    'excel_format_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/UKV-ScenariosConfig.xlsx',
    'excel_format_cfg_sheet': 'FormatConfig',
    'swimlanes_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/UKV-ScenariosConfig.xlsx',
    'swimlanes_cfg_sheet': 'Swimlanes',
    'ppt_template_file': '/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/UKV-ScenariosTemplate.pptx',
}

parameters_03 = {
    'excel_plan_file': '/Users/livestockinformation/Downloads/KBT-Plan.xlsx',
    'excel_plan_sheet': 'KBT-Plan',
    'excel_plot_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Added Value - Knowledge-Based Trading/Planning/Thomas/KBT-VisualConfig.xlsx',
    'excel_plot_cfg_sheet': 'PlotConfig',
    'excel_format_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Added Value - Knowledge-Based Trading/Planning/Thomas/KBT-VisualConfig.xlsx',
    'excel_format_cfg_sheet': 'FormatConfig',
    'swimlanes_cfg_file': '/Users/livestockinformation/Livestock Information Ltd/Added Value - Knowledge-Based Trading/Planning/Thomas/KBT-VisualConfig.xlsx',
    'swimlanes_cfg_sheet': 'Swimlanes',
    'ppt_template_file': '/Users/livestockinformation/Livestock Information Ltd/Added Value - Knowledge-Based Trading/Planning/Thomas/KBT-PlanOnePager.pptx',
}

parameters_to_use = parameters_01  # Set to whichever we are testing with or running.


def get_parameters():
    args = sys.argv
    # There should either be no parameters or 7, otherwise report error and finish
    # Note the length of argv is one more than the number of arguments as the file is always first.

    if len(args) == 1:
        return parameters_to_use

    if len(args) == 8:
        parameters = {
            'excel_plan_file': args[1],
            'excel_plan_sheet': args[2],
            'excel_plot_cfg_file': args[3],
            'excel_plot_cfg_sheet': args[4],
            'excel_format_cfg_file': args[3],
            'excel_format_cfg_sheet': args[5],
            'swimlanes_cfg_file': args[3],
            'swimlanes_cfg_sheet': args[6],
            'ppt_template_file': args[7]
        }
        return parameters

    root_logger.error(f'Wrong number of parameters provided ({len(args)-1}).  Should be 0 or 6')
    return None


def configure_logger(logger):
    log_formatter = logging.Formatter("[%(levelname)-5.5s] %(asctime)s [%(threadName)-12.12s] %(message)s")

    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)

    # Probably doesn't need to rotate files as the log file is always created each time the app is run.
    file_handler = RotatingFileHandler(f"{'/Users/livestockinformation/PycharmProjects/ppt-plan-visual/logging'}/{'plan_to_ppt'}-{time_string}.log")
    file_handler.setFormatter(log_formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    logger.setLevel(logging.INFO)


def main():
    configure_logger(root_logger)

    root_logger.debug('Plan to PowerPoint plotting programme starting...')
    root_logger.info(f"Running from IDE, using fixed arguments")

    parameters = get_parameters()
    if parameters is not None:

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
