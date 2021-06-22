class LayoutAttributes:
    def __init__(
            self,
            swimlane_name,
            track_number,
            number_of_tracks_to_span,


    ):
        """
        Encapsulates the parameters which drive how an activity is plotted within the visual.
        """
        self.swimlane_name = swimlane_name
        self.track_number = track_number
        self.number_of_tracks_to_span = number_of_tracks_to_span
