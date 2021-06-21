from measurement.measures import Distance


class PlanVisual:
    """
    Class to represent all the elements of a visualisation of a plan, but agnostic of how the visual will
    be physically realised.

    The intention is that the physical realisation of the visual can be executed by pluggable components.
    There will be a component for PowerPoint, and one for SVG at least.

    One reason for going this route is to improve usability of the software, and in particular, to make
    it easier to implement as a web app.

    A visualisation of a plan will be created by freeing plan elements from the GANTT chart format and plotting them in
    a way which makes much better use of the space available, and also using visual properties such as colours to make
    the visual representation of the plan as pleasing and helpful as possible.

    The plotting of a plan into a visual includes the following features:

    - The positioning of elements will be up to the user based on the best visual impact.  Activities could appear at
      the same line (vertical spacing) on the visual as long as they don't overlap, and this will help make better use
      of the space.

    - Configurable plotting shapes and properties: An activity will typically be plotted as a long rectangle, with the
      length representing the elapsed time that the activity is expected to take.

    - Swimlanes: Wherever an activity on the plan may appear within the Work Breakdown Structure, on the visual it can
      be placed within a given swimlane which is chosen specifically to visually group elements of the plan for which
      this makes sense. So, for examples, key stage gate milestones could be gathered together into a single
      "Governance" or similar swimlane, even though, within the plan, they may not be placed physically close to each
      other (typically they would be located within the activities that lead to the milestone being completed).

    - Hierarchical layout: A potentially powerful way of laying out elements which visually represents the hierarchical
      structure of activities, sub-activities etc by nesting visual elements within other visual elements.
    """

    def __init__(self,
                 visual_width: Distance = None,
                 visual_height: Distance = None,
                 visual_start_date: Distance = None,
                 visual_end_date: Distance = None
                 ):
        """
        When creating the visual, we may know what the dimensions need to be or we may wish them to be defined by the
        content.  Dimensions, if specified will use the Distance class from the distance packages.  This allows the
        app to then convert into whatever unit is convenient for plotting later on.

        :type visual_width: Distance
        :type visual_height: Distance
        :type visual_start_date: Distance
        :type visual_end_date: Distance

        """
        self.visual_width = visual_width
        self.visual_height = visual_height
        self.visual_start_date = visual_start_date
        self.visual_end_date = visual_end_date


