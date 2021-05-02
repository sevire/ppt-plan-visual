#!/usr/bin/env bash

# Remove pre-existing build and dist folders in case there was a problem in previous run
rm -rf build
rm -rf dist

pyinstaller source/ppt_plot_plan_main.py --onefile
mv dist/ppt_plot_plan_main .

ls -l ppt_plot_plan_main

# Remove build and dist folders after build and move
rm -rf build
rm -rf dist