from colour import Color
from measurement.measures import Distance


class VisualElementDisplayAttributes:
    """
    Encapsulates visual display attributes which are relevant for any visual element to be displayed.
    """
    def __init__(self,
                 line_thickness: Distance,
                 line_colour: Color,
                 fill_colour: Color,
                 ):
        self.line_thickness = line_thickness
        self.line_colour = line_colour
        self.fill_colour = fill_colour
