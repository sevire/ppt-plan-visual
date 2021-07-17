from typing import Union, Optional

from colour import Color
from pptx.util import Cm

from source.visualiser.plot_driver import PlotDriver
from source.visualiser.text_formatting import TextFormatting


class ShapeFormatting:
    """
    Encapsulates visual display attributes which are relevant for any visual element to be displayed.
    """
    def __init__(self,
                 line_colour: Union[Color, None],
                 fill_colour: Union[Color, None],
                 corner_radius: Union[Cm, None] = None,
                 text_formatting: Optional[TextFormatting] = None
                 ):
        self.line_colour = line_colour
        self.fill_colour = fill_colour
        self.corner_radius = corner_radius
        self.text_formatting = text_formatting

    @classmethod
    def from_dict(cls, format_dict, plot_config: PlotDriver):
        fill_colour = Color(rgb=map(lambda x: x/255, format_dict['fill_rgb']))
        line_colour = Color(rgb=map(lambda x: x/255, format_dict['line_rgb']))
        corner_radius = format_dict['corner_radius']

        margin_top = plot_config.text_margin
        margin_left = plot_config.text_margin
        margin_bottom = plot_config.text_margin
        margin_right = plot_config.text_margin
        vertical_align = format_dict['text_vertical_align']
        font_size = format_dict['font_size']
        font_bold = format_dict['font_bold']
        font_italic = format_dict['font_italic']
        font_colour = format_dict['font_colour_rgb']

        text_formatting = TextFormatting(
            margin_top=margin_top,
            margin_left=margin_left,
            margin_bottom=margin_bottom,
            margin_right=margin_right,
            vertical_align=vertical_align,
            font_size=font_size,
            font_bold=font_bold,
            font_italic=font_italic,
            font_colour=font_colour
        )

        return ShapeFormatting(
            line_colour,
            fill_colour,
            corner_radius,
            text_formatting
        )
