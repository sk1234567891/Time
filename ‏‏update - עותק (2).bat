setlocal EnableDelayedExpansion

 timeout -t 10
ping 8.8.8.8 | findstr "Reply from ?.?.?.?: bytes=32"
	IF !errorlevel!==0 (
 	REM Exit PowerPoint
cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\time2.xlsm
timeout -t 40

:EX
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
	IF !ERRORLEVEL!==0 (
	"D:\time\ex\CloseEX.vbs"
	timeout -t 3
	taskkill /IM EXCEL.EXE
	goto EX
	) 

:PP
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL
	IF !ERRORLEVEL!==0 (
	"D:\time\pp\ClosePP.vbs"
	taskkill /IM POWERPNT.EXE
	timeout -t 3
	goto PP
	) 

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