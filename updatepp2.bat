setlocal EnableDelayedExpansion

rem timeout -t 10
ping 8.8.8.8 | findstr "Reply from ?.?.?.?: bytes=32"
	IF %errorlevel%==0 (
 	REM Exit PowerPoint
 taskkill /IM POWERPNT.EXE
timeout -t 3
rem pause
 	REM Ensure PowerPoint exit
 taskkill /IM POWERPNT.EXE
cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\time2.xlsm
timeout -t 10
rem pause
	REM save the time file
"D:\time\ex\close.vbs"
	REM Exit the time file
taskkill /IM EXCEL.EXE
rem timeout -t 10
rem pause
set /p test=<D:\time\ex\holidays.txt
IF !test!==0 (
	date /t |findstr Fri
	IF !errorlevel!==0 (
						REM IF Friday this will run
							"D:\time\pp\Shabat\OpenPP.vbs"
						exit
						)
		REM IF not Friday then this will run
			"D:\time\pp\hol\OpenPP.vbs"
		exit						
	)
"D:\time\OpenPP.vbs" !test!
)
exit