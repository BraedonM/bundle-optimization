@echo off
REM Set the destination directory (relative to current directory)
set "DEST_DIR=dist\startupBundleOptimizer\_internal"

REM Copy all required files from the current directory to the destination

copy "app_icon_alt.ico" "%DEST_DIR%"
copy "help_icon.png" "%DEST_DIR%"

copy "Sub-Bundle_Data.xlsx" "%DEST_DIR%"
copy "SO_Input_Example.xlsx" "%DEST_DIR%"
copy "Packaging_Data.xlsx" "%DEST_DIR%"

copy "variables.json" "%DEST_DIR%"

echo Files copied to %DEST_DIR%.
pause

