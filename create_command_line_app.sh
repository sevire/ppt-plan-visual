#!/usr/bin/env bash

# Remove pre-existing build and dist folders in case there was a problem in previous run
rm -rf build
rm -rf dist

echo "Build and Dist folders removed. Press a key to continue"
read dummy

pyinstaller source/ppt_plot_plan_main.py --onefile
mv dist/ppt_plot_plan_main .

ls -l ppt_plot_plan_main

echo "Command line has been built.  Press a key to remove Build and Dist folders."
read dummy

# Remove build and dist folders after build and move
rm -rf build
rm -rf dist