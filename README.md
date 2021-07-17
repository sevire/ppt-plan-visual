# PowerPoint Plan Visual - Overview.

Python app to generate a PowerPoint visualisation of a project.

The app is designed to help in solving the problem that many Project Managers and similar business professionals have to solve all the time, which is that of how to create a visual representation of a complex and detailed project plan on a single page, making use of all the space on the page and choosing the layout of the visual so that it makes sense to stakeholders who need to understand a plan at a high level.

# What The App Does

1. Takes an Excel representation of a plan, where there is one row for each activity or milestone.  This can either be hand crafted or an extract from a planning tool such as SmartSheets.
2. Uses information provided by the user in additional columns to determine how to layout activities and milestones from the plan.
3. Reads in a template PowerPoint file which is used as the starting point for the visual.
4. The visual is then created by adding shapes to the first slide of the template file.
5. The app adds a shape for each activity which has been flagged by the user as being included in the visual.
6. Each shape added is place horizontally on the slide based on the start date and end date.
7. Each shape is placed vertically on the screen by specifying a named swimlane for the activity and a vertical "track" within that swimlane.
8. The width of each activity on the visual will be determined by the start and end dates.
9. The height of each activity is specified by the user, and defaults to 1 track.
10. The text description/title of each activity is plotted separately from the shape itself, and can be positioned in a number of layout options driven mainly by whether the shape is large enough to contain the text.
11. The layout options are:
    - Shape: To coincide with the activity itself.
    - Left: The text is placed so that it covers the shape but extends to the left of the shape.
    - Right: The text is places so that it covers the shape but extends to the right of the shape.
    (Other options will be added)
10. The shapes corresponding to each activity or milestone can be formatted according to a set of formatting options provided by the user, which specify a number of formatting characteristics, including:
    - Fill colour
    - Line colour
    - Font size
    - Bold
    - Italics
    - (various others)

# Additional Usage Notes

If the originating plan has a hierarchical work breakdown structure comprising high level activities with lower level activities, then the use can choose whether or not to include both summary tasks and detailed tasks, and also how to lay these out.  In particular if the requirement is to place a larger (higher) shape for summary tasks, and place sub-tasks within the larger task, then this can be done using the layout options (swimlane + vertical track + activity height).

