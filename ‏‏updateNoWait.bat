setlocal EnableDelayedExpansion

 timeout -t 10
ping 8.8.8.8 | findstr "Reply from ?.?.?.?: bytes=32"
	IF %errorlevel%==0 (
 	REM Exit PowerPoint
cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\time2.xlsm
timeout -t 40
"D:\time\ex\CloseEX.vbs"
"D:\time\pp\ClosePP.vbs"
 timeout -t 3
 
set /p test=<D:\time\ex\holidays.txt
IF !test!==0 (
	date /t |findstr -c:"Fri" -c:"Sat"
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