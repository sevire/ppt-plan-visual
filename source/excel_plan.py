import logging

import pandas as pd
import numpy as np

root_logger = logging.getLogger()


class ExcelPlan:
    def __init__(self, excel_driver_config, excel_plan_file):
        self.excel_plan_sheet_name = excel_driver_config['excel_plan_sheet_name']
        self.excel_plot_config_sheet_name = excel_driver_config['excel_plot_config_sheet_name']
        self.excel_format_config_sheet_name = excel_driver_config['excel_format_config_sheet_name']
        self.plan_start_row = excel_driver_config['plan_start_row']

        self.xl_pd_object = pd.ExcelFile(excel_plan_file)

    def read_plan_data(self):
        milestones = self.xl_pd_object.parse(self.excel_plan_sheet_name, skiprows=self.plan_start_row - 1)
        milestones.set_index('Id', inplace=True)

        plan_data = []

        for row_id, milestone_data in milestones.iterrows():
            # Will probably need to pre-process dates so readable by Python
            start_date = milestone_data['Start Date']
            end_date = milestone_data['End Date']

            record = {
                'id': row_id,
                'description': milestone_data['Description'],
                'type': milestone_data['Activity Type'],
                'start_date': start_date,
                'end_date': end_date,
                'swimlane': milestone_data['Swimlane'],
                'track_num': milestone_data['Visual Track Number Within Swimlane'],
                'bar_height_in_tracks': milestone_data['Visual Num Tracks To Span'],
                'format_properties': milestone_data['Format Name']
            }
            plan_data.append(record)

        return plan_data


class ExcelSmartsheetPlan:
    """
    Excel import but specifically customised for the SmartSheet plan format used in a real plan.
    Main difference is the column names and the fact that not all rows are to be included.

    The rows which are to be included have the column "Visual Flag" set to true
    """
    def __init__(self, excel_driver_config, excel_plan_file):
        self.excel_plan_sheet_name = excel_driver_config['excel_plan_sheet_name']
        self.excel_plot_config_sheet_name = excel_driver_config['excel_plot_config_sheet_name']
        self.excel_format_config_sheet_name = excel_driver_config['excel_format_config_sheet_name']
        self.plan_start_row = excel_driver_config['plan_start_row']

        read_cols =[
            'Task Name',
            'Visual Text',
            'Duration',
            'Start',
            'Finish',
            'Visual Flag',
            'Visual Swimlane',
            'Visual Track # Within Swimlane',
            'Visual # Tracks To Cover',
            'Text Layout',
            'Format String',
            'Done Format String'
        ]

        converters = {
            'Visual Flag': ExcelSmartsheetPlan.bool_converter
        }

        pd_object_with_nan = pd.read_excel(
            excel_plan_file,
            engine='openpyxl',
            sheet_name=self.excel_plan_sheet_name,
            usecols=read_cols,
            converters=converters
        )
        self.xl_pd_object = pd_object_with_nan
        # ToDo: Decide how we are going to deal with NaNs

    def read_plan_data(self):
        milestones = self.xl_pd_object
        # milestones = self.xl_pd_object.parse(self.excel_plan_sheet_name, skiprows=self.plan_start_row - 1)
        # milestones.set_index('Id', inplace=True)

        plan_data = []

        for index, milestone_data in milestones.iterrows():
            flag = milestone_data['Visual Flag']
            if flag is True:
                start_date = milestone_data['Start']
                end_date = milestone_data['Finish']
                duration = milestone_data['Duration']
                description = milestone_data['Task Name']
                visual_text = milestone_data['Visual Text']
                visual_swimlane = milestone_data['Visual Swimlane']
                track_num = milestone_data['Visual Track # Within Swimlane']
                num_tracks = milestone_data['Visual # Tracks To Cover']
                format_properties = milestone_data['Format String']

                # Pre-processing and setting defaults for missing values

                if pd.isnull(visual_text):
                    text = description
                else:
                    text = visual_text

                if duration == '0':
                    activity_type = 'milestone'
                    root_logger.debug(f'Activity [{description}:40.40] is a milestone')
                else:
                    activity_type = 'bar'
                    root_logger.debug(f'Activity [{description}:40.40] is an activity')

                if pd.isnull(visual_swimlane):
                    root_logger.warning(f'No swimlane specified for [{description:40.40}], setting to "Default"')
                    visual_swimlane = 'Default'

                if pd.isnull(track_num):
                    root_logger.warning(f'No track num specified for [{description:40.40}], setting to 1')
                    track_num = 1

                if pd.isnull(num_tracks):
                    root_logger.warning(f'Num tracks not specified for [{description:40.40}], setting to 1')
                    num_tracks = 1

                if pd.isnull(format_properties):
                    root_logger.warning(f'Format name not specific for [{description:40.40}], setting to "Default"')
                    format_properties = 'Default'

                record = {
                    'id': index,
                    'description': text,
                    'type': activity_type,
                    'start_date': start_date,
                    'end_date': end_date,
                    'swimlane': visual_swimlane,
                    'track_num': track_num,
                    'bar_height_in_tracks': num_tracks,
                    'format_properties': format_properties,
                    'done_format_properties': milestone_data['Done Format String'],
                    'text_layout': milestone_data['Text Layout']
                }
                plan_data.append(record)

        return plan_data

    @staticmethod
    def bool_converter(smartsheet_flag_value):
        if smartsheet_flag_value is True:
            return True
        else:
            return False
