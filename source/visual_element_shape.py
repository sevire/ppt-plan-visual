from enum import Enum


class VisualElementShape(Enum):
    """
    Represents the shape attributes for a specific shape appropriate for a plotable element.

    Example, rectangle, rounded_rectangle.
    """
    RECTANGLE = 1
    ROUNDED_RECTANGLE = 2