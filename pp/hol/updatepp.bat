timeout -t 10
ping 8.8.8.8 | findstr "Reply from ?.?.?.?: bytes=32"
	IF %errorlevel%==0 (
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
"D:\time\pp\hol\OpenPP.vbs"
)
exit