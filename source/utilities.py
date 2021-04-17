import os


def get_path_name_ext(path):
    folder = os.path.dirname(path)
    file = os.path.basename(path)
    base, ext = os.path.splitext(file)

    return folder, base, ext


class SwimlaneManager:
    """
    - Manages list of swimlanes which drives positioning of swimlanes on visual.
    - Takes list of swimlanes and uses order in list to determine swimlane number.
    - Implements method to return swimlane number to user during visual creation.
      If a request for the number of a non-existent swimlane is made, the class will add the swimlane to the end of the list
      and return it's implied number, to ensure consistency.
    """

    def __init__(self, swimlane_data):
        self.swimlane_data = swimlane_data

    def get_swimlane_number(self, swimlane_name):
        # Make sure there is an entry for this swimlane.  If not already there, add it.
        if swimlane_name not in self.swimlane_data:
            self.swimlane_data.append(swimlane_name)

        # Filter out all entries which aren't the indicated swimlane.
        # If there is a duplicate, ignore and just return the lowest.  That will be consistent.
        this_swimlane_only = [index for index, swimlane in enumerate(self.swimlane_data) if swimlane == swimlane_name]
        return this_swimlane_only[0] + 1
