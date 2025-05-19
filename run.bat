@echo off
REM This batch file runs the Python script "run_automation.py"
REM It first changes the current directory to the location of this batch file.

REM Change directory to the script's own directory
REM %~dp0 expands to the drive and path of the current batch script
echo Changing directory to script location: %~dp0
cd /d %~dp0

REM --- Option 1: Python is in PATH ---
REM This command assumes that the 'python' command is accessible
REM through your system's PATH environment variable.
echo Starting Python script 'run_automation.py' from current directory...
python run_automation.py

REM --- Option 2: Specify full path to Python executable ---
REM If 'python' is not in your PATH, or you need to use a specific
REM Python installation, uncomment the line below and replace
REM "C:\Path\To\Your\Python\python.exe" with the actual full path
REM to your python.exe.
REM "C:\Path\To\Your\Python\python.exe" run_automation.py

REM Keep the command window open after the script finishes (optional)
REM If you want the command prompt to stay open after the script
REM has finished (e.g., to see output or errors), uncomment the next line.
REM pause

echo Script execution finished.
