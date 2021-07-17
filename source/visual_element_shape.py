from enum import Enum
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE


class VisualElementShape(Enum):
    """
    Represents the shape attributes for a specific shape appropriate for a plotable element.
    Also includes mapping to the actual shape type to use to plot on a PPT slide.

    Example, rectangle, rounded_rectangle.
    """
    RECTANGLE = (1, MSO_AUTO_SHAPE_TYPE.RECTANGLE)
    ROUNDED_RECTANGLE = (2, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE)
    DIAMOND = (3, MSO_AUTO_SHAPE_TYPE.DIAMOND)

    def __init__(self, index, ppt_shape):
        self.index = index
        self.ppt_shape = ppt_shape

    def ppt_shape(self):
        return self.ppt_shape

