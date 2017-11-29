set currentpath=%~dp0
set PYTHONPATH=%PYTHONPATH%;%currentpath%\Lib;
python handle_excel.py
pause