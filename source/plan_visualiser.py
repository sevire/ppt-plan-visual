import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PP_ALIGN
from pptx.enum.text import MSO_VERTICAL_ANCHOR as MSO_ANCHOR
from source.excel_plan import ExcelPlan, ExcelSmartsheetPlan
from source.plot_driver import PlotDriver
from source.utilities import get_path_name_ext, SwimlaneManager


class PlanVisualiser:
    """
    Object which manages the creation of a PPT slide with a visual representation
    of a project plan (or similar).

    The supplied data

    :param plan_data:
    """

    def __init__(self, plan_data, plot_config, format_config, template_path, slide_level_config):
        # The actual plan data with activities and milestones, start/finish dates etc.
        self.plan_data = plan_data

        # Data to define plot area for elements, tracks etc.
        self.plot_config = plot_config

        # Data with pre-determined formatting properties to apply to elements.
        self.format_config = format_config

        # Config to drive slide level objects such as the swimlane rectangle shapes as background.
        self.slide_level_config = slide_level_config

        self.template = template_path

        folder, base, ext = get_path_name_ext(template_path)
        self.slides_out_path = os.path.join(folder, base + '_outx' + ext)
        self.prs = Presentation(template_path)

        visual_slide = self.prs.slides[0]  # Assume there is one slide and that's where we will place the visual

        self.shapes = visual_slide.shapes
        self.plot_driver = PlotDriver(plot_config)

        self.swimlane_data = self.extract_swimlane_data()

    @classmethod
    def from_excel(cls, plan_data_excel_path, plot_area_config, format_config, excel_driver_config, template_path, slide_level_config):
        excel_manager = ExcelPlan(excel_driver_config, plan_data_excel_path)

        extracted_plot_config = plot_area_config
        extracted_format_config = format_config
        extracted_plan_data = excel_manager.read_plan_data()

        return PlanVisualiser(extracted_plan_data, extracted_plot_config, extracted_format_config, template_path, slide_level_config)

    @classmethod
    def from_excelsmartsheet(cls, plan_data_excel_path, plot_area_config, format_config, excel_driver_config, template_path, slide_level_config):
        excel_manager = ExcelSmartsheetPlan(excel_driver_config, plan_data_excel_path)

        extracted_plot_config = plot_area_config
        extracted_format_config = format_config
        extracted_plan_data = excel_manager.read_plan_data()

        return PlanVisualiser(extracted_plan_data, extracted_plot_config, extracted_format_config, template_path, slide_level_config)

    def plot_slide(self):
        """
        Opens a supplied template file in order to allow consistency with other slides in a deck.

        Then creates a new slide with a plotted plan visual based on the data read in to the object.

        Then writes the one-slide deck to a different filename in the same folder.

        :return:
        """

        self.plot_swimlanes(self.slide_level_config)
        for plotable_element in self.plan_data:
            start = plotable_element['start_date']
            end = plotable_element['end_date']
            description = plotable_element['description']
            swimlane = plotable_element['swimlane']
            track_num = plotable_element['track_num']
            num_tracks = plotable_element['bar_height_in_tracks']
            shape_format_name = plotable_element['format_properties']
            format_data = self.format_config['format_categories'][shape_format_name]
            text_layout = plotable_element['text_layout']
            if plotable_element['type'] == 'bar':
                self.plot_bar(description, start, end, swimlane, track_num, num_tracks, format_data, text_layout)

            elif plotable_element['type'] == 'milestone':
                self.plot_milestone(description, start, swimlane, track_num, format_data, text_layout)

        self.prs.save(self.slides_out_path)

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
        swimlane_manager = SwimlaneManager(self.slide_level_config['swimlanes'])
        for record in self.plan_data:
            swimlane = record['swimlane']
            track_num_high = record['track_num'] + record['bar_height_in_tracks'] - 1
            if swimlane not in swimlane_data:
                # Adding record for this swimlane. Also remember ordering as is important later.
                swimlane_record = {
                    'swimlane_number': swimlane_manager.get_swimlane_number(swimlane),
                    'highest_track_within_lane': track_num_high
                }
                swimlane_data[swimlane] = swimlane_record
            else:
                swimlane_record = swimlane_data[swimlane]
                if track_num_high > swimlane_record['highest_track_within_lane']:
                    swimlane_record['highest_track_within_lane'] = track_num_high

        # We now have a dict containing a record for each swimlane of lowest and highest relative track number used.
        # Can now calculate the start and end track number for each swimlane - in order that lanes were encountered.

        swimlane_plot_data = {}
        end_track = 0

        # Create list of swimlane data entries in order of swimlane number
        sorted_entries = sorted(swimlane_data.items(), key=lambda x: x[1]['swimlane_number'])

        # Calculate start and end track number for each swimlane based on number of tracks required for swimlane
        for swimlane, swimlane_entry in sorted_entries:
            start_track = end_track + 1
            end_track = start_track + swimlane_entry['highest_track_within_lane'] - 1
            swimlane_plot_data[swimlane] = {
                'start_track': start_track,
                'end_track': end_track
            }
        return swimlane_plot_data

    def plot_bar(self, activity_description, start_date, end_date, swimlane, track_number, num_tracks,
                 properties, text_layout):
        swimlane_start = self.swimlane_data[swimlane]['start_track']

        left = self.plot_driver.date_to_x_coordinate(start_date)
        right = self.plot_driver.date_to_x_coordinate(end_date)
        width = right - left

        top = self.plot_driver.track_number_to_y_coordinate(swimlane_start + track_number - 1)
        height = self.plot_driver.height_of_track(num_tracks)

        shape = self.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height
        )

        # Adjust rounded corner radius
        target_radius = properties['corner_radius']
        adjustment_value = target_radius / height

        shape.adjustments[0] = adjustment_value

        self.shape_fill(shape, properties)
        self.shape_line(shape, properties)

        activity_text_width = self.plot_config['activity_text_width']
        if text_layout == 'Left':
            # Extend the text to the left so that if overflows to the left of the shape.
            adjust_width = max(width, activity_text_width)
            text_left = left + width - adjust_width
            text_width = activity_text_width
            properties['text_align'] = 'right'
            self.plot_text(activity_description, text_left, top, text_width, height, properties)
        elif text_layout == 'Right':
            # Extend text to the right so that it overflows to the right of the shape
            adjust_width = max(width, activity_text_width)
            text_left = left
            text_width = adjust_width
            properties['text_align'] = 'left'
            self.plot_text(activity_description, text_left, top, text_width, height, properties)
        else:  # Apply default which is "Shape"
            # Standard positioning, text will align exactly with the shape
            properties['text_align'] = 'centre'
            self.plot_text(activity_description, left, top, width, height, properties)

    def plot_milestone(self, milestone_description, start_date, swimlane, track_number, properties, text_layout):
        swimlane_start = self.swimlane_data[swimlane]['start_track']
        milestone_width = self.plot_config['milestone_width']
        milestone_height = self.plot_config['track_height']

        left = self.plot_driver.date_to_x_coordinate(start_date) - milestone_width / 2
        top = self.plot_driver.track_number_to_y_coordinate(swimlane_start + track_number - 1)

        shape = self.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.DIAMOND, left, top, milestone_width, milestone_height
        )

        self.shape_fill(shape, properties)
        self.shape_line(shape, properties)

        if text_layout == "Right":
            milestone_text_width = self.plot_config['milestone_text_width']
            milestone_text_left = left + milestone_width
            properties['text_align'] = 'left'
            self.plot_text(milestone_description, milestone_text_left, top, milestone_text_width, milestone_height, properties)
        else:  # Default is left
            milestone_text_width = self.plot_config['milestone_text_width']
            milestone_text_left = left - milestone_text_width
            properties['text_align'] = 'right'
            self.plot_text(milestone_description, milestone_text_left, top, milestone_text_width, milestone_height, properties)

    def plot_text(self, text, left, top, width, height, format_data):

        shape = self.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height
        )
        shape.fill.background()
        shape.line.fill.background()

        self.add_text_to_shape(shape, text, format_data)

    def add_text_to_shape(self, shape, text, format_data):
        text_frame = shape.text_frame
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0

        vertical = format_data['text_vertical_align']
        if vertical == "top":
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
        elif vertical == "bottom":
            text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        else:
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        paragraph = text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = text

        self.text_format(paragraph, run, format_data)

    def shape_fill(self, shape, format_data):
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*format_data['fill_rgb'])

    def text_format(self, paragraph, run, format_data):
        """
        Hard code text formatting (for now) as is likely to be constant.

        :param paragraph:
        :param run:
        :return:
        """
        font = run.font
        paragraph.line_spacing = 0.8
        paragraph.alignment = self.text_alignment(format_data['text_align'])

        font.name = 'Calibri'
        font.size = format_data['font_size']
        font.bold = format_data['font_bold']
        font.italic = format_data['font_italic']
        font.color.rgb = RGBColor(*format_data['font_colour_rgb'])

    @staticmethod
    def shape_line(shape, format_data):
        line = shape.line
        line.color.rgb = RGBColor(*format_data['line_rgb'])

    @staticmethod
    def text_alignment(format_text_align):
        """
        Takes the alignment field from the element formatting data and converts to the appropriate value for the
        pptx-python setting.

        :param format_text_align:
        :return:
        """
        if format_text_align == "left":
            return PP_ALIGN.LEFT
        elif format_text_align == "right":
            return PP_ALIGN.RIGHT
        elif format_text_align == "centre":
            return PP_ALIGN.CENTER
        else:
            # Default to centre
            return PP_ALIGN.CENTER

    def plot_swimlanes(self, format_data):
        """
        plot a background rectangle for each swimline with the width of the plot area, the height of the swimlane, and
        positioned to coincide with the tracks within the swimlane.

        Also alternate colours for each swimlane

        :return:
        """

        for row, swimlane in enumerate(self.swimlane_data):
            row_number = row + 1

            start_track = self.swimlane_data[swimlane]['start_track']
            end_track = self.swimlane_data[swimlane]['end_track']
            top = self.plot_driver.track_number_to_y_coordinate(start_track)

            # Need to adjust to start between track end and track start, unless this is the first row
            if row_number > 1:
                top -= (self.plot_config['track_gap'] / 2)
            bottom = self.plot_driver.track_number_to_y_coordinate(end_track) + self.plot_config['track_height'] + (self.plot_config['track_gap'] / 2)
            left = self.plot_config['left']
            width = self.plot_config['right'] - self.plot_config['left']
            height = bottom - top

            shape = self.shapes.add_shape(
                MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height
            )

            # For the purposes of this decision, the first row is 1 (odd)
            if row_number % 2 == 0:
                format_info = format_data['swimlane_format_even']
            else:
                format_info = format_data['swimlane_format_odd']

            self.shape_fill(shape, format_info)
            self.shape_line(shape, format_info)
            self.add_text_to_shape(shape, swimlane, format_info)
