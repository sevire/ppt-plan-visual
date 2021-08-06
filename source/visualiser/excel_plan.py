import logging

from source.visualiser.exceptions import PptPlanVisualiserException
from source.visualiser.activity_layout_attributes import ActivityLayoutAttributes
from source.visualiser.plan_activity import PlanActivity
from source.visualiser.read_excel import read_excel
from source.visualiser.shape_formatting import ShapeFormatting

root_logger = logging.getLogger()


class ExcelPlan:
    """
    Excel import but specifically customised for the SmartSheet plan format used in a real plan.
    Main difference is the column names and the fact that not all rows are to be included.

    The rows which are to be included have the column "Visual Flag" set to true
    """

    @staticmethod
    def read_plan_data(
            excel_plan_file,
            excel_plan_sheet_name,
            format_properties_list,
            plan_visual_config
    ):

        read_cols =[
            'Task Name',
            'Visual Text',
            'Duration',
            'Start',
            'Finish',
            'Visual Flag',
            'Visual Swimlane',
            'Visual Track # Within Swimlane',
            'Visual # Tracks To Cover',
            'Text Layout',
            'Format String',
            'Done Format String'
        ]

        converters = {
            'Visual Flag': ExcelPlan.bool_converter
        }

        pd_object = read_excel(excel_plan_file, excel_plan_sheet_name)
        plan_data = []
        swimlane_max_track_num = {}

        for index, milestone_data in enumerate(pd_object):
            flag = milestone_data['Visual Flag']
            if flag is True:
                start_date = milestone_data['Start']
                end_date = milestone_data['Finish']
                duration = milestone_data['Duration']
                description = milestone_data['Task Name']
                visual_text = milestone_data['Visual Text']
                visual_swimlane = milestone_data['Visual Swimlane']
                track_num = milestone_data['Visual Track # Within Swimlane']
                num_tracks = milestone_data['Visual # Tracks To Cover']
                format_1_id = milestone_data['Format String']
                format_2_id = milestone_data['Done Format String']
                text_layout = milestone_data['Text Layout']

                # Pre-processing and setting defaults for missing values

                if visual_text is None:
                    text = description
                else:
                    text = visual_text

                if duration == '0' or duration == 0:
                    activity_type = 'milestone'
                    root_logger.debug(f'Activity [{description}:40.40] is a milestone')
                else:
                    activity_type = 'bar'
                    root_logger.debug(f'Activity [{description}:40.40] is an activity')

                if visual_swimlane is None:
                    root_logger.warning(f'No swimlane specified for [{description:40.40}], setting to "Default"')
                    visual_swimlane = 'Default'

                # Allocate track num if not already set and update max track num for this swimlane if necessary.

                # Remember that track_num is None so that log message can be output later.
                if track_num is None:
                    if visual_swimlane not in swimlane_max_track_num:
                        track_num = 1
                        swimlane_max_track_num[visual_swimlane] = 1
                    else:
                        track_num = swimlane_max_track_num[visual_swimlane] + 1
                        swimlane_max_track_num[visual_swimlane] = max(track_num, swimlane_max_track_num[visual_swimlane])
                    root_logger.warning(f'No track num specified for [{description:40.40}], setting to {track_num}')
                else:
                    if visual_swimlane not in swimlane_max_track_num:
                        swimlane_max_track_num[visual_swimlane] = track_num
                    swimlane_max_track_num[visual_swimlane] = max(track_num, swimlane_max_track_num[visual_swimlane])

                if num_tracks is None:
                    root_logger.warning(f'Num tracks not specified for [{description:40.40}], setting to 1')
                    num_tracks = 1

                if format_1_id is None:
                    root_logger.warning(f'Format name not specified for [{description:40.40}], setting to "Default"')
                    format_1_id = 'Default'

                if format_2_id is None:
                    root_logger.warning(f'Format name not specified for [{description:40.40}], setting to "Default"')
                    format_2_id = None

                if text_layout is None:
                    # Text layout isn't specified, so:
                    # - if it's a milestone position to left
                    # - if it's an activity, position within shape

                    if activity_type == "milestone":
                        root_logger.warning(f'Text layout for {activity_type} not specified for [{description:40.40}], setting to "Left"')
                        text_layout = 'Left'
                    elif activity_type == "bar":
                        root_logger.warning(f'Text layout for {activity_type} not specific for [{description:40.40}], setting to "Shape"')
                        text_layout = 'Left'
                    else:
                        root_logger.warning(f'Text layout not specific for [{description:40.40}], setting to "Left"')
                        raise PptPlanVisualiserException(f'Unknown value for activity_type ({activity_type})')

                activity_layout_attributes = ActivityLayoutAttributes(
                    visual_swimlane,
                    track_num,
                    num_tracks,
                    text_layout,
                )

                shape_formatting_1 = ShapeFormatting.from_dict(format_properties_list[format_1_id], plan_visual_config)
                if format_2_id is None:
                    shape_formatting_2 = None
                else:
                    shape_formatting_2 = ShapeFormatting.from_dict(format_properties_list[format_2_id], plan_visual_config)

                if activity_type == "bar":
                    display_shape = plan_visual_config.activity_shape
                elif activity_type == "milestone":
                    display_shape = plan_visual_config.milestone_shape
                else:
                    raise PptPlanVisualiserException(f"Unexpected activity type '{activity_type}'")

                activity = PlanActivity(
                    activity_id=index,
                    description=text,
                    activity_type=activity_type,
                    start_date=start_date,
                    end_date=end_date,
                    activity_layout_attributes=activity_layout_attributes,
                    display_shape=display_shape,
                    plan_visual_config=plan_visual_config,
                    shape_formatting_1=shape_formatting_1,
                    shape_formatting_2=shape_formatting_2
                )

                plan_data.append(activity)

        return plan_data

    @staticmethod
    def bool_converter(smartsheet_flag_value):
        if smartsheet_flag_value is True:
            return True
        else:
            return False
