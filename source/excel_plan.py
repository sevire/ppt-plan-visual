import pandas as pd


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

        self.xl_pd_object = pd.ExcelFile(excel_plan_file)

    def read_plan_data(self):
        milestones = self.xl_pd_object.parse(self.excel_plan_sheet_name, skiprows=self.plan_start_row - 1)
        # milestones.set_index('Id', inplace=True)

        plan_data = []

        for index, milestone_data in milestones.iterrows():
            flag = milestone_data['Visual Flag']
            if flag == 1:
                start_date = milestone_data['Start']
                end_date = milestone_data['Finish']
                duration = milestone_data['Duration']
                description = milestone_data['Task Name']
                visual_text = milestone_data['Visual Text']
                if pd.isnull(visual_text):
                    text = description
                else:
                    text = visual_text
                if duration == '0':
                    activity_type = 'milestone'
                else:
                    activity_type = 'bar'

                record = {
                    'id': index,
                    'description': text,
                    'type': activity_type,
                    'start_date': start_date,
                    'end_date': end_date,
                    'swimlane': milestone_data['Visual Swimlane'],
                    'track_num': milestone_data['Visual Track # Within Swimlane'],
                    'bar_height_in_tracks': milestone_data['Visual # Tracks To Cover'],
                    'format_properties': milestone_data['Format String']
                }
                plan_data.append(record)

        return plan_data
