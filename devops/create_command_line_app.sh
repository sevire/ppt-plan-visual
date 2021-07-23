#!/usr/bin/env bash

# Remove pre-existing build and dist folders in case there was a problem in previous run
rm -rf app_build
rm -rf app_dist

pyinstaller source/ppt_plot_plan_main.py --onefile
mv app_dist/ppt_plot_plan_main .

ls -l ppt_plot_plan_main

# Remove build and dist folders after build and move
rm -rf app_build
rm -rf app_dist