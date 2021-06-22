from datetime import date
from source.common_display_attributes import VisualElementDisplayAttributes
from source.exceptions import PptPlanVisualiserException
from source.layout_attributes import LayoutAttributes
from source.plot_driver import PlotDriver
from source.visual_element_shape import VisualElementShape


class PlanActivity:
    """
    The class represents a plotable object which is to be placed on the plan visual.  The plotable object is an
    activity, which represents a period of time to be plotted.

    An activity will be plotted onto the visual canvas as a rectangle typically, but there will be some flexibility to
    allow compelling visual displays where this is appropriate.

    The laying out of the shape will be specified separately from the details of the activity.

    An activity has:
    - A start date
    - An end date
    - Display attributes; colours, shape to use etc.

    There will be the option of considering the activity in two parts, the 'done' bit, and the 'still to do' bit.  This
    is to cater for the use case where the user wants to (typically) display the done portion of the activity in
    a different colour.
    """
    def __init__(
            self,
            start_date: date,
            end_date: date,
            layout_attributes: LayoutAttributes,
            display_attributes: VisualElementDisplayAttributes,
            display_shape: VisualElementShape,
            visual_config: PlotDriver,
            swimlane_data: dict,
            done_display_attributes: VisualElementDisplayAttributes = None
    ):
        """
        :param start_date: Start date for this activity
        :param end_date: End date for this activity
        :param layout_attributes: Drives where activity is plotted vertically and how high, etc.
        :param display_attributes: Drives various formatting attributes for the activity such as line colour,
               fill colour etc.
        :param display_shape:
        :param visual_config:
        :param done_display_attributes: Drives formatting attributes for the 'done' part of the activity if user wants
               to include it.  Absence of this parameter means don't split into done and not done.
        """
        self.start_date = start_date
        self.end_date = end_date
        self.layout_attributes = layout_attributes
        self.display_attributes = display_attributes
        self.done_display_attributes = done_display_attributes
        self.display_shape = display_shape
        self.swimlane_data = swimlane_data  # ToDo: Incorporate swimlane data into PlotConfig (or maybe plan data)
        self.visual_config = visual_config
        self.today = self.visual_config.today

    def is_current(self):
        return self.start_date < self.today < self.end_date

    def is_past(self):
        return self.start_date < self.end_date <= self.today

    def is_future(self):
        return self.today < self.start_date < self.end_date

    @property
    def include_done(self):
        if self.done_display_attributes is None:
            return False
        else:
            return True

    @property
    def plot_start_x(self):
        """
        Return the x coordinate to plot for the start date of an activity (or the to-do bit of an
        activity). Various cases depending upon whether the activity is to be split or not and whether
        it is in the past, is current, or is in the future.

        :return:
        """
        if self.include_done is True:
            if self.is_past():
                return None  # There is no to-do bit so nothing to plot.
            elif self.is_current():
                # This is the to-do bit of a current activity so start is today.
                return self.visual_config.date_to_x_coordinate(self.today)
            elif self.is_future():
                # This is the to-do bit of a future activity so start is the activity start.
                return self.visual_config.date_to_x_coordinate(self.start_date)
            else:
                raise PptPlanVisualiserException('Error in processing date in done/to do logic')
        else:
            # We aren't splitting into done and to-do so just return the actual activity start.
            return self.visual_config.date_to_x_coordinate(self.start_date)

    @property
    def plot_end_x(self):
        if self.include_done:
            if self.is_past():
                return None
            else:
                return self.visual_config.date_to_x_coordinate(self.end_date)
        else:
            return self.visual_config.date_to_x_coordinate(self.end_date)

    @property
    def plot_done_start_x(self):
        if self.include_done is True:
            if self.is_future():
                return None
            else:
                return self.visual_config.date_to_x_coordinate(self.start_date)
        else:
            return None

    @property
    def plot_done_end_x(self):
        if self.include_done is True:
            if self.is_past():
                return self.visual_config.date_to_x_coordinate(self.end_date)
            elif self.is_current():
                return self.visual_config.date_to_x_coordinate(self.visual_config.today)
            elif self.is_future():
                return None
            else:
                raise PptPlanVisualiserException('Error in processing date in done/to do logic')
        else:
            return None

    @property
    def plot_width(self):
        if self.plot_end_x is None or self.plot_start_x is None:
            return None
        else:
            return self.plot_end_x - self.plot_start_x

    @property
    def plot_done_width(self):
        if self.include_done:
            if self.plot_done_end_x is None or self.plot_done_start_x is None:
                return None
            else:
                return self.plot_done_end_x - self.plot_done_start_x
        else:
            return None

    @property
    def plot_top(self):
        swimlane_start = self.swimlane_data[self.layout_attributes.swimlane_name]['start_track']
        track_number = self.layout_attributes.track_number
        top = self.visual_config.track_number_to_y_coordinate(swimlane_start + track_number - 1)

        return top

    @property
    def plot_height(self):
        return self.visual_config.height_of_track(self.layout_attributes.number_of_tracks_to_span)

    def get_ppt_plot_coords(self, done_flag=False):
        """
        Returns parameters to be passed to the PPT plot_shape method.
        :param done_flag: If set and the include_done flag is set, return the 'done' coords, otherwise
                          return main activity coordinates.
        :return:
        """
        if done_flag is True:
            if self.include_done:
                ret = (self.plot_done_start_x, self.plot_top, self.plot_done_width, self.plot_height)
                if None in ret:
                    return None
                else:
                    return ret
            else:
                return None
        else:
            ret = (self.plot_start_x, self.plot_top, self.plot_width, self.plot_height)
            if None in ret:
                return None
            else:
                return ret
