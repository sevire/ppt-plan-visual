from dataclasses import dataclass
from datetime import date
from typing import Union

from colour import Color
from pptx.util import Cm, Pt

from source.visualiser.plotable_element import PlotableElement
from source.visualiser.shape_formatting import ShapeFormatting
from source.visualiser.exceptions import PptPlanVisualiserException
from source.visualiser.activity_layout_attributes import ActivityLayoutAttributes
from source.visualiser.plot_driver import PlotDriver
from source.visualiser.text_formatting import TextFormatting
from source.visualiser.visual_element_shape import VisualElementShape


@dataclass
class PlanActivity:
    """
    The class represents a plotable object which is to be placed on the plan visual.  The plotable object is an
    activity, which represents a period of time to be plotted.

    The activity may be a milestone which has zero duration and is plotted as a single point in time rather than as a
    period.

    An Activity will be plotted onto the visual canvas as a rectangle typically, but the user can specify any one of
    a number of supported shapes.  Similarly a Milestone will typically be plotted as a Diamond, but other possibilities
    exist which can be selected by the user.

    There is the ability to distinguish between the past and the future, and plot an activity split into two separate
    parts, with the past being formatted differently (usually just by fill colour).  For Milestones, rather than plot as
    two separate parts, only one shape will be plotted, but it will be plotted with different formatting if it is in the
    past.

    Given the above, there are a number of different cases and associated calculations of plot points to consider.

    Key drivers for plotting calculations:
    - Is this an activity or a milestone?
    - Is multiple formatting enabled for this activity?
    - Is the activity wholly in the past, current (straddles current date) or wholly in the future.

    format_1 will be used to describe the 'main' formatting which applies in the following cases:
    - If multiple formatting is not enabled.
    - For the whole activity or milestone if it is wholly in the future.
    - For the future part of an activity which straddles the current date.

    format_2 will be used to describe the secondary formatting which applies in the following cases:
    - Multiple formatting is enabled.
    - For the whole activity or milestone if it is wholly in the past.
    - For the past part of the activity (not milestone) if it straddles the current date.

    NOTE: The text associated with the activity will be plotted in exactly the same way regardless of whether muliple
    formatting is enabled, and will be placed as though the activity was plotted as a single entity, and use the
    text formatting supplied with format_1

    An activity has:
    - A start date
    - An end date
    - Display attributes; colours, shape to use etc.

    start_date: Start date for this activity
    end_date: End date for this activity
    layout_attributes: Drives where activity is plotted vertically and how high, etc.
    display_attributes: Drives various formatting attributes for the activity such as line colour,
    fill colour etc.
    display_shape:
    plot_visual_config:
    done_display_attributes: Drives formatting attributes for the 'done' part of the activity if user wants
    to include it.  Absence of this parameter means don't split into done and not done.
    """
    activity_id: int
    description: str
    activity_type: str
    start_date: date
    end_date: date
    activity_layout_attributes: ActivityLayoutAttributes
    display_shape: VisualElementShape
    plan_visual_config: PlotDriver
    shape_formatting_1: ShapeFormatting
    shape_formatting_2: Union[ShapeFormatting, None] = None
    today_override: Union[date, None] = None
    swimlane_start_track: Union[int, None] = None

    @property
    def today(self):
        if self.today_override is not None:
            return self.today_override
        else:
            return self.plan_visual_config.today

    @property
    def text_formatting(self):
        """
        The formatting parameters to use will be taken from formatting_1.

        :return:
        """
        return self.shape_formatting_1.text_formatting

    @property
    def multi_format_enabled(self):
        if self.shape_formatting_2 is None:
            return False
        else:
            return True

    def ppt_plot_shape(
            self,
            ppt_shapes_object,
            display_shape,
            top,
            left,
            width,
            height,
            shape_formatting,
            text=None,
    ):
        plot_element = PlotableElement(
            shape=display_shape,
            top=top,
            left=left,
            bottom=top + height,
            right=left + width,
            shape_formatting=shape_formatting,
            text=text,
            text_formatting=self.text_formatting,
        )
        shape = plot_element.plot_ppt(ppt_shapes_object)
        return shape

    def plot_ppt_shapes(self, ppt_shapes_object):
        """
        Works out what to plot and plots it on a PowerPoint slide (supplied)

        :param ppt_shapes_object:
        :return:
        """
        shapes = []  # Collect all shapes plotted and return them
        if not self.multi_format_enabled:
            # Simple case.  Just plot one activity shape and one text shape with formatting_1
            if self.activity_type == "milestone":
                left, top, width, height = self.get_milestone_coords()
                shape = self.ppt_plot_shape(
                    ppt_shapes_object,
                    self.display_shape,
                    top,
                    left,
                    width,
                    height,
                    self.shape_formatting_1
                )
                shapes.append(shape)
            elif self.activity_type == "bar":
                left = self._shape_left("activity")
                top = self._plot_top
                width = self._shape_width("activity")
                height = self._plot_height
                shape = self.ppt_plot_shape(
                    ppt_shapes_object,
                    self.display_shape,
                    top=top,
                    left=left,
                    width=width,
                    height=height,
                    shape_formatting=self.shape_formatting_1
                )
                shapes.append(shape)
            else:
                raise PptPlanVisualiserException(f"Unexpected activity type '{self.activity_type}'")
        else:
            # Multiple formats are enabled so work out which case
            if self.activity_type == "milestone":
                # No splitting required as it's a milestone, but work out whether the milestone is
                # in the past and use alternative formatting if it is.
                if self.is_past():
                    formatting = self.shape_formatting_2
                else:
                    formatting = self.shape_formatting_1
                left = self._shape_left("milestone")
                top = self._plot_top
                width = self._shape_width("milestone")
                height = self._plot_height
                shape = self.ppt_plot_shape(
                    ppt_shapes_object,
                    self.display_shape,
                    top,
                    left,
                    width,
                    height,
                    formatting
                )
                shapes.append(shape)
            else:
                # Multiple formats for an activity (not a milestone).  There are three cases.
                if self.is_past() or self.is_future():
                    # We are only plotting one shape, with alternative formatting if in past.
                    formatting = self.shape_formatting_2 if self.is_past() else self.shape_formatting_1
                    left = self._shape_left("activity")
                    top = self._plot_top
                    width = self._shape_width("activity")
                    height = self._plot_height
                    shape = self.ppt_plot_shape(
                        ppt_shapes_object,
                        self.display_shape,
                        top,
                        left,
                        width,
                        height,
                        formatting
                    )
                    shapes.append(shape)
                else:
                    # Most complex case.  We are plotting the activity as two shapes, the past and the future.
                    # The past has the alternative formatting, the future has the default formatting.

                    top = self._plot_top
                    height = self._plot_height

                    # Calculate 'past' shape values
                    left_1 = self._shape_left("activity")
                    width_1 = self._shape_width("part_1")

                    # Calculate 'future' shape values
                    left_2 = self._shape_left("today")
                    width_2 = self._shape_width("part_2")

                    shape = self.ppt_plot_shape(
                        ppt_shapes_object,
                        self.display_shape,
                        top,
                        left_1,
                        width_1,
                        height,
                        self.shape_formatting_2
                    )
                    shapes.append(shape)

                    shape = self.ppt_plot_shape(
                        ppt_shapes_object,
                        self.display_shape,
                        top,
                        left_2,
                        width_2,
                        height,
                        self.shape_formatting_1
                    )
                    shapes.append(shape)
        shape = self.plot_ppt_text_shape(ppt_shapes_object)
        shapes.append(shape)
        return shapes

    def is_current(self):
        if self.start_date <= self.today <= self.end_date:
            return True
        else:
            return False

    def is_past(self):
        return self.start_date < self.today and self.end_date < self.today

    def is_future(self):
        return self.start_date > self.today and self.end_date > self.today

    def _shape_left(self, case) -> int:
        """
        returns correct left plot value for an activity/milestone shape for the given case.  Cases are:
        - activity: Return value corresponding to the start date of the activity
        - today: Return left value corresponding to the current date
        - milestone: return left value of milestone shape based on start date of milestone

        :return: left value to be used in ppt plot.
        """
        if case == "activity":
            return self.plan_visual_config.date_to_x_coordinate(self.start_date, "start")
        elif case == "today":
            return self.plan_visual_config.date_to_x_coordinate(self.today, "start")
        elif case == "milestone":
            milestone_width = self.plan_visual_config.milestone_width
            milestone_left_adjust = milestone_width / 2

            # The milestone symbol needs to be plotted so that it's centre coincides exactly with the
            # date it represents.  So offset the left point by half the width of the symbol.
            start_date_left = self.plan_visual_config.date_to_x_coordinate(self.start_date, "middle")
            return round(start_date_left - milestone_left_adjust)
        else:
            raise PptPlanVisualiserException(f"Unexpected case supplied '{case}'while calculating left plot")

    def _shape_width(self, case) -> int:
        """
        returns correct width value for the given case.  Cases are:
        - activity: Return width of whole activity
        - part_1: Return width of start date to today's date
        - part_2: Return width of today's date to end date
        - milestone: Return width of milestone symbol to plot

        :return: width value to be used in ppt plot.
        """

        if case == "activity":
            end = self.plan_visual_config.date_to_x_coordinate(self.end_date, "end")
            start = self.plan_visual_config.date_to_x_coordinate(self.start_date)
            return round(end - start)
        elif case == "part_1":
            end = self.plan_visual_config.date_to_x_coordinate(self.today, "start")
            start = self.plan_visual_config.date_to_x_coordinate(self.start_date)
            return round(end - start)
        elif case == "part_2":
            end = self.plan_visual_config.date_to_x_coordinate(self.end_date, "end")
            start = self.plan_visual_config.date_to_x_coordinate(self.today)
            return round(end - start)
        elif case == "milestone":
            return round(self.plan_visual_config.milestone_width)

    @property
    def _plot_top(self):
        if self.swimlane_start_track is None:
            swimlane_name = self.activity_layout_attributes.swimlane_name
            raise PptPlanVisualiserException(f"Swimlane start track for '{swimlane_name}' not available while calculating top of activity")
        track_number = self.swimlane_start_track + self.activity_layout_attributes.track_number - 1
        top = self.plan_visual_config.track_number_to_y_coordinate(track_number)

        return round(top)

    @property
    def _plot_height(self):
        return round(self.plan_visual_config.height_of_track(self.activity_layout_attributes.number_of_tracks_to_span))

    def get_ppt_text_coords(self):
        """
        Works out where to place text for this activity given layout attributes.
        :return:
        """
        text_bottom = self._plot_top + self._plot_height
        left = self._shape_left("activity")
        text_top = self._plot_top
        width = self._shape_width("activity")

        min_activity_text_width = max(width, self.plan_visual_config.min_activity_text_width)
        if self.activity_layout_attributes.text_layout == 'Left':
            # Extend the text to the left so that if overflows to the left of the shape.
            adjust_width = max(width, min_activity_text_width)
            text_left = left + width - adjust_width
            text_right = text_left + adjust_width
            text_align = 'right'
        elif self.activity_layout_attributes.text_layout == 'Right':
            # Extend text to the right so that it overflows to the right of the shape
            adjust_width = max(width, min_activity_text_width)
            text_left = left
            text_right = text_left + adjust_width
            text_align = 'left'
        else:  # Apply default which is "Shape"
            # Standard positioning, text will align exactly with the shape
            text_align = 'centre'
            text_left = left
            text_right = text_left + width
        return text_top, text_left, text_bottom, text_right, text_align

    def plot_ppt_text_shape(self, ppt_shapes_object):
        text_top, text_left, text_bottom, text_right, text_align = self.get_ppt_text_coords()

        left_margin = self.plan_visual_config.text_margin
        right_margin = self.plan_visual_config.text_margin
        margin_adjust = round(self.plan_visual_config.milestone_width / 2) + self.plan_visual_config.text_margin
        if self.activity_type == "milestone":
            # Need to add margin to move text outside milestone shape
            if text_align == "right":
                right_margin = margin_adjust
            elif text_align == "left":
                left_margin = margin_adjust
        if self.text_formatting is None:
            # Provide default text_formatting, but shouldn't really happen!

            text_formatting = TextFormatting(
                margin_top=Cm(0),
                margin_left=left_margin,
                margin_bottom=Cm(0),
                margin_right=right_margin,
                vertical_align='middle',
                horizontal_align=text_align,
                font_size=Pt(8),
                font_bold=False,
                font_italic=False,
                font_colour=Color(rgb=(0, 0, 0))
            )
        else:
            text_formatting = self.text_formatting

        text_formatting.horizontal_align = text_align
        text_formatting.margin_left = left_margin
        text_formatting.margin_right = right_margin

        shape_formatting = ShapeFormatting(
            line_colour=None,
            fill_colour=None,
            corner_radius=None,
            text_formatting=text_formatting
        )
        plot_element = PlotableElement(
            shape=VisualElementShape.RECTANGLE,
            top=text_top,
            left=text_left,
            bottom=text_bottom,
            right=text_right,
            shape_formatting=shape_formatting,
            text=self.description,
            text_formatting=text_formatting
        )
        shape = plot_element.plot_ppt(ppt_shapes_object)
        return shape

    def get_milestone_coords(self):
        left = self._shape_left("milestone")
        top = self._plot_top
        width = self._shape_width("milestone")
        height = self._plot_height
        return left, top, width, height

    def __str__(self):

        if self.is_current():
            status = 'current'
        elif self.is_future():
            status = 'future'
        elif self.is_past():
            status = 'past'
        else:
            status = 'error'
        return f'[{self.start_date}:{self.end_date}, {status}]'
