import logging
import pandas as pd

root_logger = logging.getLogger()


class ExcelPlan:
    """
    Excel import but specifically customised for the SmartSheet plan format used in a real plan.
    Main difference is the column names and the fact that not all rows are to be included.

    The rows which are to be included have the column "Visual Flag" set to true
    """

    @staticmethod
    def read_plan_data(excel_plan_file, excel_plan_sheet_name):

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
            'Visual Flag': ExcelPlan.bool_converter
        }

        pd_object = pd.read_excel(
            excel_plan_file,
            engine='openpyxl',
            sheet_name=excel_plan_sheet_name,
            usecols=read_cols,
            converters=converters
        )
        plan_data = []

        for index, milestone_data in pd_object.iterrows():
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
                text_layout = milestone_data['Text Layout']

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

                if pd.isnull(text_layout):
                    # Text layout isn't specified, so:
                    # - if it's a milestone position to left
                    # - if it's an activity, position within shape

                    if activity_type == "milestone":
                        root_logger.warning(f'Text layout for {activity_type} not specified for [{description:40.40}], setting to "Left"')
                        text_layout = 'Left'
                    elif activity_type == "bar":
                        root_logger.warning(f'Text layout for {activity_type} not specific for [{description:40.40}], setting to "Shape"')
                        text_layout = 'Shape'
                    else:
                        root_logger.warning(f'Text layout not specific for [{description:40.40}], setting to "Left"')
                        raise Exception(f'Unknown value for activity_type ({activity_type})')


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
                    'text_layout': text_layout
                }
                plan_data.append(record)

        return plan_data

    @staticmethod
    def bool_converter(smartsheet_flag_value):
        if smartsheet_flag_value is True:
            return True
        else:
            return False
