class ActivityLayoutAttributes:
    """
    Encapsulates the parameters which drive the layout of a specific activity or milestone.
    """
    def __init__(
            self,
            swimlane_name,
            track_number,
            number_of_tracks_to_span,
            text_layout
    ):
        self.swimlane_name = swimlane_name
        self.track_number = track_number
        self.number_of_tracks_to_span = number_of_tracks_to_span
        self.text_layout = text_layout

