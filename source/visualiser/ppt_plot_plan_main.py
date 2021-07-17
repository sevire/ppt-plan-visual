import logging
import sys
import time
from logging.handlers import RotatingFileHandler
from source.visualiser.plan_visualiser import PlanVisualiser

root_logger = logging.getLogger()

# No args provided so use hard-coded defaults

parameters_01 = {
    'excel_plan_file': '/Users/thomasdeveloper/Documents/Projects/ppt-plan-visual-data/PlanningVisualConfig-01a.xlsx',
    'excel_plan_sheet': 'UK-View Plan',
    'excel_config_workbook': '/Users/thomasdeveloper/Documents/Projects/ppt-plan-visual-data/PlanningVisualConfig-01a.xlsx',
    'ppt_template_file': '/Users/thomasdeveloper/Documents/Projects/ppt-plan-visual-data/UK-ViewPlanOnePager.pptx',
}
parameters_02 = {
    'excel_plan_file': '~/Downloads/KBT-Delivery.xlsx',
    'excel_plan_sheet': 'KBT-Delivery',
    'excel_config_workbook': '~/PyCharmProjects/ppt_plan_visual_testing/KBT-VisualConfig.xlsx',
    'ppt_template_file': '/Users/Development/PycharmProjects/ppt_plan_visual_testing/KBT-DeliveryOnePager.pptx',
}

parameters_to_use = parameters_02  # Set to whichever we are testing with or running.


def get_parameters():
    """
    Gets command line parameters.  There should be 4 parameters which are:
    - Excel Plan File
    - Excel Config File
    - Excel Plan Sheet Name: Defaults to the name of the file as that is what is used in SmartSheets
    - PPT Template File: Takes first slide as template for output.

    :return:
    """
    args = sys.argv
    # There should either be no parameters or 7, otherwise report error and finish
    # Note the length of argv is one more than the number of arguments as the file is always first.

    if len(args) == 1:
        return parameters_to_use

    if len(args) == 8:
        parameters = {
            'excel_plan_workbook': args[1],
            'excel_plan_sheet': args[2],
            'excel_config_workbook': args[3],
            'ppt_template_file': args[4],
        }
        return parameters

    root_logger.error(f'Wrong number of parameters provided ({len(args)-1}).  Should be 0 or 4')
    return None


def configure_logger(logger):
    log_formatter = logging.Formatter("[%(levelname)-5.5s] %(asctime)s [%(threadName)-12.12s] %(message)s")

    ts = time.gmtime()
    time_string = time.strftime("%Y-%m-%d_%H:%M:%S", ts)

    # Probably doesn't need to rotate files as the log file is always created each time the app is run.
    file_handler = RotatingFileHandler(f"{'plan_to_ppt'}-{time_string}.log")
    file_handler.setFormatter(log_formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    logger.setLevel(logging.DEBUG)


def main():
    configure_logger(root_logger)
    parameters = get_parameters()
    if parameters is not None:
        excel_plan_file = parameters['excel_plan_file']
        excel_plan_sheet = parameters['excel_plan_sheet']
        excel_config_workbook = parameters['excel_config_workbook']
        ppt_template_file = parameters['ppt_template_file']

        visualiser = PlanVisualiser.from_excel(excel_plan_file, excel_config_workbook, ppt_template_file, excel_plan_sheet)
        visualiser.plot_slide()


if __name__ == '__main__':
    main()
