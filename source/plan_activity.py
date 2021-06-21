from datetime import date
from source.common_display_attributes import VisualElementDisplayAttributes


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
    """
    def __init__(self,
                 start_date: date,
                 end_date: date,
                 display_attributes: VisualElementDisplayAttributes,
                 display_shape: VisualElementShape
                 ):
        self.start_date = start_date
        self.end_date = end_date
        self.display_attributes = display_attributes
