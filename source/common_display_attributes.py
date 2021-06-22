from colour import Color


class VisualElementDisplayAttributes:
    """
    Encapsulates visual display attributes which are relevant for any visual element to be displayed.
    """
    def __init__(self,
                 line_colour: Color,
                 fill_colour: Color,
                 font_colour: Color
                 ):
        self.line_colour = line_colour
        self.fill_colour = fill_colour
        self.font_colour = font_colour

    @classmethod
    def from_dict(cls, format_dict):
        fill_colour = Color(rgb=map(lambda x: x/255, format_dict['fill_rgb']))
        line_colour = Color(rgb=map(lambda x: x/255, format_dict['line_rgb']))
        font_colour = Color(rgb=map(lambda x: x/255, format_dict['font_colour_rgb']))

        return VisualElementDisplayAttributes(line_colour, fill_colour, font_colour)
