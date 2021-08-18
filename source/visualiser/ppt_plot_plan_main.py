import logging
import sys
import time
from logging.handlers import RotatingFileHandler
from source.visualiser.plan_visualiser import PlanVisualiser

root_logger = logging.getLogger()

# No args provided so use hard-coded defaults

parameters_01 = {
    'excel_plan_workbook': '/Users/Development/Downloads/UK-View-Delivery.xlsx',
    'excel_plan_sheet': 'UK-View-Delivery',
    'excel_config_workbook': '/Users/Development/PycharmProjects/ppt_plan_visual_testing/KBT-VisualConfig.xlsx',
    'ppt_template_file': '/Users/Development/PycharmProjects/ppt_plan_visual_testing/UKView-Del-DeliveryOnePager.pptx',
}
parameters_02 = {
    'excel_plan_workbook': '/Users/Development/Downloads/KBT-Delivery.xlsx',
    'excel_plan_sheet': 'KBT-Delivery',
    'excel_config_workbook': '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/config_files/KBT-VisualConfig.xlsx',
    'ppt_template_file': '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/input_files/ppt_templates/PlanVisual-01.pptx',
}

parameters_03 = {
    'excel_plan_workbook': '/Users/Development/CommandLine/scripts/BethPhDGANTT-visual.xlsx',
    'excel_plan_sheet': 'Sheet1',
    'excel_config_workbook': '/Users/Development/CommandLine/scripts/Beth-PHD-GANTT-config.xlsx',
    'ppt_template_file': '/Users/Development/CommandLine/scripts/Beth-PHD-GANTT.pptx',
}

parameters_04 = {
    'excel_plan_workbook': '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/unit_test_01/input_files/unit_test_01_config.xlsx',
    'excel_plan_sheet': 'Plan',
    'excel_config_workbook': '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/unit_test_01/input_files/unit_test_01_config.xlsx',
    'ppt_template_file': '/Users/Development/PycharmProjects/ppt-plan-visual/source/tests/test_resources/unit_test_01/input_files/unit_test_dummy_ppt.pptx',
}


parameters_to_use = parameters_04  # Set to whichever we are testing with or running.


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

    expected_num_args = 4
    if len(args) == 1:
        return parameters_to_use

    if len(args) == expected_num_args + 1:
        parameters = {
            'excel_plan_workbook': args[1],
            'excel_plan_sheet': args[2],
            'excel_config_workbook': args[3],
            'ppt_template_file': args[4],
        }
        return parameters

    root_logger.error(f'Wrong number of parameters provided ({len(args)-1}).  Should be 0 or {expected_num_args}')
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
        excel_plan_file = parameters['excel_plan_workbook']
        excel_plan_sheet = parameters['excel_plan_sheet']
        excel_config_workbook = parameters['excel_config_workbook']
        ppt_template_file = parameters['ppt_template_file']

        visualiser = PlanVisualiser.from_excel(excel_plan_file, excel_config_workbook, ppt_template_file, excel_plan_sheet)
        visualiser.plot_slide()


if __name__ == '__main__':
    main()
