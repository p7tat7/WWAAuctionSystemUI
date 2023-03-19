@echo off
cd src


:: run the python script
"C:\Users\p7tat7\AppData\Local\Programs\Python\Python39\python.exe" "C:\Users\p7tat7\World Wide Aquarium Dropbox\WWA SHKS\PROGRAM\AuctionSystem UI\src\main.py"


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