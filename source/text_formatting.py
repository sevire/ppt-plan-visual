from dataclasses import dataclass

from colour import Color
from pptx.util import Cm, Pt


@dataclass
class TextFormatting:
    """
    margin_top:
    margin_left:
    margin_bottom:
    margin_right:
    vertical_align: 'top', 'middle', 'bottom'
    """
    margin_top: Cm = 0
    margin_left: Cm = 0
    margin_bottom: Cm = 0
    margin_right: Cm = 0
    vertical_align: str = 'middle'
    horizontal_align: str = 'centre'
    font_size: Pt = Pt(10)
    font_bold: bool = False
    font_italic: bool = False
    font_colour: Color = Color(rgb=(0.1, 0.1, 0.1))

