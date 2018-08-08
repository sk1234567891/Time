timeout -t 10
	REM Exit PowerPoint slide show
start /wait cscript "D:\time\ExitSlideShow.vbs"
timeout -t 3
	REM Exit PowerPoint
taskkill /IM POWERPNT.EXE
timeout -t 3
	REM Ensure PowerPoint exit
taskkill /IM POWERPNT.EXE
cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\time2.xlsm
timeout -t 60
	REM save the time file
"D:\time\ex\close.vbs"
	REM Exit the time file
taskkill /IM EXCEL.EXE
cd /D D:\time
	REM Enter the GUI file
start POWERPNT.EXE D:\time\pp\Shabat\show2.pptx
timeout /t 3
	REM Take focus on the windows
"D:\time\PPAppFocus.vbs"
	REM Navigate to the right button
tabbutton.vbs
timeout /t 3
tabbutton.vbs
timeout /t 3
enterbutton.vbs
timeout -t 2
	REM Wait for PowerPoint to finish update from excel
	
:loop
	tasklist | findstr EXCEL.EXE
	if %errorlevel%==0 (
						timeout -t 3
						goto loop
						)
						
	REM Saving the changes so in the next update it go smooth
start /wait cscript "D:\time\pp\Shabat\save.vbs"
	REM Entering the slide show and ensure it enter the slide show
f5button.vbs
timeout /t 60
f5button.vbs