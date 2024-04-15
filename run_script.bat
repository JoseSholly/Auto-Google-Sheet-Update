@echo off

rem Set the path to your virtual environment activation script
set VENV_PATH=C:\Users\USER\Documents\Python.py\Automation\Auto_spread_sheet\venv\Scripts\activate.bat

rem Activate the virtual environment
call "%VENV_PATH%"

rem Run your Python script within the virtual environment
python C:\Users\USER\Documents\Python.py\Automation\Auto_spread_sheet/sheet_update.py


