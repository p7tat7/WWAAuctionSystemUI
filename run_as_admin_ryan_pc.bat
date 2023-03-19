@echo off
cd src


:: run the python script
"C:\Users\user\AppData\Local\Programs\Python\Python310\python.exe" "C:\Users\user\World Wide Aquarium Dropbox\WWA SHKS\PROGRAM\AuctionSystem UI\src\main.py"


:: check if there is error during the script
if %errorlevel% == 0 (goto endnormally)



echo.
echo.
echo There is error during the program, please contact Ryan/Tommy and provide the log file.
pause

:: exit program
exit

:endnormally
:: exit program
exit