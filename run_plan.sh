#!/usr/bin/env bash

source /Users/livestockinformation/PycharmProjects/ppt-plan-visual/venv/bin/activate
python -m source.ppt_plot_plan_main \
'/Users/livestockinformation/Downloads/UK-View Plan.xlsx' \
'UK-View Plan' \
'/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/PlanningVisualConfig-01.xlsx' \
'PlotConfig' \
'FormatConfig' \
'Swimlanes' \
'/Users/livestockinformation/Livestock Information Ltd/Data - UK data/UK View/planning/planning-visual/UK-ViewPlanOnePager.pptx'
