import os
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from datetime import datetime
import pandas as pd

from pptx.util import Cm

from source.tests.test_data.test_data_01 import plot_area_config, format_config
from source.utilities import get_path_name_ext


class PlanVisualiser:
    """
    Object which manages the creation of a PPT slide with a visual representation
    of a project plan (or similar).

    The supplied data

    :param plan_data:
    """

    def __init__(self, plan_data, plot_config, format_config, template_path):
        # The actual plan data with activities and milestones, start/finish dates etc.
        self.plan_data = plan_data

        # Data to define plot area for elements, tracks etc.
        self.plot_config = plot_config

        # Data with pre-determined formatting properties to apply to elements.
        self.format_config = format_config

        self.template = template_path

        folder, base, ext = get_path_name_ext(template_path)
        self.slides_out_path = os.path.join(folder, base + '_out' + ext)
        self.prs = Presentation(template_path)

        visual_slide = self.prs.slides[0]  # Assume there is one slide and that's where we will place the visual

        self.shapes = visual_slide.shapes
        self.plot_driver = PlotDriver(plot_area_config)

        self.swimlane_driver = self.extract_swimlane_data()

    @classmethod
    def from_excel(cls, plan_data_excel_path, excel_driver_config, template_path):
        excel_manager = ExcelPlan(excel_driver_config, plan_data_excel_path)

        extracted_plot_config = excel_manager.extract_plot_config_data()
        extracted_format_config = excel_manager.extract_format_config_data()
        extracted_plan_data = excel_manager.read_plan_data()

        return PlanVisualiser(extracted_plan_data, extracted_plot_config, extracted_format_config, template_path)

    def extract_swimlane_data(self):
        """
        After some thought am taking a very simple approach here.

        Just go through each activity on the plan and:
        - Check which swimlane it is for
        - Look at the track number and height in tracks and therefore work out the bottom track number
        - If the bottom track number is larger than that already recorded, update to reflect.
        - Once you have checked every activity, the entry in each swimlane will have recorded the number of tracks
          required for that swimlane.
        - We can then go back and calculate the start and end track for each swimlane which is what the plot method
          needs.

          So we will end up with a dict, with one entry for each (named) swimlane, and against each swimlane we will
          see the start track number and the end track number.

        :return:
        """

        swimlane_data = {}
        entry_num = 0
        for record in self.plan_data:
            swimlane = record['swimlane']
            track_num_low = record['track_num']
            track_num_high = record['track_num'] + record['bar_height_in_tracks'] - 1
            if swimlane not in swimlane_data:
                # Adding record for this swimlane. Also remember ordering as is important later.
                swimlane_record = {
                    'swimlane_order': entry_num + 1,
                    'lowest_track_within_lane': track_num_low,
                    'highest_track_within_lane': track_num_high
                }
                swimlane_data[swimlane] = swimlane_record
                entry_num += 1
            else:
                swimlane_record = swimlane_data[swimlane]
                if track_num_low < swimlane_record['lowest_track_within_lane']:
                    swimlane_record['lowest_track_within_lane'] = track_num_low
                if track_num_high > swimlane_record['highest_track_within_lane']:
                    swimlane_record['highest_track_within_lane'] = track_num_high

        # We now have a dict containing a record for each swimlane of lowest and highest relative track number used.
        # Can now calculate the start and end track number for each swimlane - in order that lanes were encountered.

        swimlane_plot_data = {}
        end_track = 0
        for lane_number in range(0, len(swimlane_data)):
            swimlane_entries = [(key, swimlane_data[key]) for key in swimlane_data.keys() if swimlane_data[key]['swimlane_order'] == lane_number + 1]

            assert(len(swimlane_entries) == 1)

            swimlane, swimlane_entry = swimlane_entries[0]
            tracks_for_swimlane = \
                swimlane_entry['highest_track_within_lane'] - swimlane_entry['lowest_track_within_lane'] + 1
            start_track = end_track + 1
            end_track = end_track + tracks_for_swimlane - 1
            swimlane_plot_data[swimlane] = {
                'start_track': start_track,
                'end_track': end_track
            }
        return swimlane_plot_data

    def plot_slide(self):
        """
        Opens a supplied template file in order to allow consistency with other slides in a deck.

        Then creates a new slide with a plotted plan visual based on the data read in to the object.

        Then writes the one-slide deck to a different filename in the same folder.

        :param template_path:
        :return:
        """

        for plotable_element in self.plan_data:
            if plotable_element['type'] == 'bar':
                start = plotable_element['start_date']
                end = plotable_element['end_date']
                swimlane = plotable_element['swimlane']
                track_num = plotable_element['track_num']
                num_tracks = plotable_element['bar_height_in_tracks']
                shape_format = plotable_element['format_properties']

                # start_date = self.plot_driver.string_date_to_date(start)
                # end_date = self.plot_driver.string_date_to_date(end)

                self.plot_bar(start, end, swimlane, track_num, num_tracks, shape_format)

        self.prs.save(self.slides_out_path)

    def plot_bar(self, start_date, end_date, swimlane, track_number, num_tracks, format):
        swimlane_start = self.swimlane_driver[swimlane]['start_track']

        left = self.plot_driver.date_to_x_coordinate(start_date)
        right = self.plot_driver.date_to_x_coordinate(end_date)

        top = self.plot_driver.track_number_to_y_coordinate(swimlane_start + track_number - 1) - 1
        bottom = top + self.plot_driver.height_of_track(num_tracks)

        shape = self.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, right-left, bottom-top
        )


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
                'type': 'bar',
                'start_date': start_date,
                'end_date': end_date,
                'track_num': milestone_data['Visual Track Number Within Swimlane'],
                'bar_height_in_tracks': milestone_data['Visual Num Tracks To Span'],
                'format_properties': 1
            }
            plan_data.append(record)

        return plan_data

    def extract_plot_config_data(self):
        # Hard code during development
        return plot_area_config

    def extract_format_config_data(self):
        # Hard code during development
        return format_config


class PlotDriver:
    """
    Used to translate plan type data into data which can be used to plot shapes.

    Example - convert from a date to a horizontal position on the slide.

    """
    def __init__(self, plot_config):
        self.top = plot_config['top']
        self.left = plot_config['left']
        self.bottom = plot_config['bottom']
        self.right = plot_config['right']
        self.track_height = plot_config['track_height']
        self.track_gap = plot_config['track_gap']
        self.min_start_date = plot_config['min_start_date']
        self.max_end_date = plot_config['max_end_date']

        self.num_days_in_date_range = self.max_end_date.toordinal() - self.min_start_date.toordinal()
        self.plot_area_width = self.right - self.left

    def date_to_x_coordinate(self, date):
        """
        This isn't quite as easy as it seems if we do it to pinpoint accuracy as the fact that
        different months have different numbers of days will come into play.

        For now just calculate based on how many days through the range the date is and
        calculate distance as a proportion of total plot width.

        In fact this approach would be completely accurate if we were using days as the unit
        not months.

        :param date:
        :return:
        """

        num_days_from_min_date = date.toordinal() - self.min_start_date.toordinal()
        proportion_of_plot_width = num_days_from_min_date / self.num_days_in_date_range
        distance_from_left_of_plot_area = proportion_of_plot_width * self.plot_area_width
        x_coord = self.left + distance_from_left_of_plot_area

        return x_coord

    def track_number_to_y_coordinate(self, track_num):
        return self.top + (track_num - 1) * (self.track_height + self.track_gap)

    def height_of_track(self, num_tracks):
        return num_tracks * self.track_height + (num_tracks-1) * self.track_gap

    def string_date_to_date(self, string_date):
        return datetime.strptime(string_date, "%Y%m%d").date()


class PlotableElement:
    def __init__(self, element):
        self.element = element

    def plot_element(self, slide_object):
        pass


class Bar(PlotableElement):
    def __init__(self, start_date, end_date, bar_height, format_properties, element):
        super().__init__(element)
        self.start_date = start_date
        self.end_date = end_date
        self.bar_height = bar_height
        self.format_properties = format_properties
