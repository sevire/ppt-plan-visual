import pandas as pd
import numpy as np
from pptx.util import Cm, Pt

from source.visualiser.plot_driver import PlotDriver


class ExcelFormatConfig:
    """
    Class to read configuration records from a sheet in an Excel File
    """
    def __init__(self, excel_path, excel_sheet, skip_rows=0):
        records_with_nan = pd.read_excel(excel_path, sheet_name= excel_sheet, engine='openpyxl', skiprows=skip_rows)
        self.records = records_with_nan.replace(np.nan, None)

    def parse_format_config(self):
        format_config_records = {}
        for id, format_excel_record in self.records.iterrows():
            format_name = format_excel_record['Format Name']
            fill_red = format_excel_record['Fill Red']
            fill_green = format_excel_record['Fill Green']
            fill_blue = format_excel_record['Fill Blue']
            line_red = format_excel_record['Line Red']
            line_green = format_excel_record['Line Green']
            line_blue = format_excel_record['Line Blue']
            font_red = format_excel_record['Font Red']
            font_green = format_excel_record['Font Green']
            font_blue = format_excel_record['Font Blue']
            config_record = {
                'fill_rgb': (fill_red, fill_green, fill_blue),
                'line_rgb': (line_red, line_green, line_blue),
                'corner_radius': Cm(format_excel_record['Corner Radius (Cm)']),
                'font_size': Pt(format_excel_record['Font Size (Pt)']),
                'font_bold': format_excel_record['Font Bold'],
                'font_italic': format_excel_record['Font Italic'],
                'font_colour_rgb': (font_red, font_green, font_blue),
                'text_vertical_align': format_excel_record['Text Vertical Align']
            }
            format_config_records[format_name] = config_record

        # Check whether a 'Default' style has been include, if not add a default default :-)
        if 'Default' not in format_config_records:
            default_record = {
                'fill_rgb': (0, 255, 255),
                'line_rgb': (255, 0, 0),
                'corner_radius': 0,
                'font_size': Pt(8),
                'font_bold': False,
                'font_italic': False,
                'font_colour_rgb': (0, 0, 0),
                'text_vertical_align': 'middle'
            }

            format_config_records['Default'] = default_record

        return format_config_records


class ExcelPlotConfig:
    """
    Class to read configuration records from a sheet in an Excel File
    """

    def __init__(self, excel_path, excel_sheet, skip_rows=0):
        records_with_nan = pd.read_excel(excel_path, sheet_name=excel_sheet, engine='openpyxl', skiprows=skip_rows)
        self.records = records_with_nan.replace(np.nan, None)

    def parse_plot_config(self):
        record = self.records.iloc[0]

        plot_area_config = {
            'top': Cm(record['Top']),
            'left': Cm(record['Left']),
            'bottom': Cm(record['Bottom']),
            'right': Cm(record['Right']),
            'track_height': Cm(record['Track Height']),
            'track_gap': Cm(record['Track Gap']),
            'min_start_date': record['Min Date'],
            'max_end_date': record['Max Date'],
            'milestone_width': Cm(record['Milestone Width']),
            'milestone_text_width': Cm(record['Milestone Text Width']),
            'activity_text_width': Cm(record['Activity Text Width']),
            'text_margin': Cm(record['Text Margin']),
            'activity_shape': record['Activity Shape'],
            'milestone_shape': record['Milestone Shape']
        }
        return PlotDriver(plot_area_config)


class ExcelSwimlaneConfig:
    """
    Class to read swimlane order
    """

    def __init__(self, excel_path, excel_sheet, skip_rows=0):
        self.records = pd.read_excel(excel_path, sheet_name=excel_sheet, engine='openpyxl', skiprows=skip_rows)

    def parse_swimlane_config(self):
        swimlanes = []
        for id, swimlane_record in self.records.iterrows():
            swimlanes.append(swimlane_record['Swimlane'])
        return swimlanes
