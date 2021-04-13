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

    def extract_plot_config_data(self):
        # Hard code during development
        return plot_area_config

    def extract_format_config_data(self):
        # Hard code during development
        return format_config