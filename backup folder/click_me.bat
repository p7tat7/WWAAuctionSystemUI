@echo off
cd src


:: run the python script
python -u main.py


:: check if there is error during the script
if %errorlevel% == 0 (goto endnormally)



echo.
echo.
echo There is error during the program, please contact Ryan/Tommy and provide the log file.
pause >nul

:: exit program
exit

:endnormally
:: exit program
exit