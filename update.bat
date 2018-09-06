setlocal EnableDelayedExpansion

 timeout -t 10
ping 8.8.8.8 | findstr "Reply from ?.?.?.?: bytes=32"
	IF !errorlevel!==0 (
 	REM Exit PowerPoint
cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\time2.xlsm 
timeout -t 40

REM ####### AutoPic function ############
REM ####### set CurSID = Current user SID
for /f "usebackq" %%i in (`wmic useraccount where "name='%username%'" get sid ^| findstr "S-"`) do (set CurSID=%%i)
REM ####### set SpotVer = Spotlight version
for /f "usebackq" %%j in (`reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Creative\!CurSID!\ ^| findstr "131"`) do (set SpotVer=%%j)
REM ####### find the spotlight current background
FOR /F "usebackq tokens=3*" %%A IN (`reg query !SpotVer! /v landscapeImage`) DO ( set BackDir=%%A )
copy !BackDir! d:\time\pp\hol\AutoPic.jpg

:PP
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>NUL | find /I /N "POWERPNT.EXE">NUL
	IF !ERRORLEVEL!==0 (
	"D:\time\pp\ClosePP.vbs"
	taskkill /IM POWERPNT.EXE
	timeout -t 10
	goto PP
	) 

set /p test=<D:\time\ex\holidays.txt
IF !test!==0 (
	date /t |findstr -c:"Fri" -c:"Sat"
	IF !errorlevel!==0 (
		REM IF Friday this will run
		"D:\time\pp\Shabat\OpenPP.vbs"
		goto EX
		exit
	)
	REM IF not Friday then this will run
	"D:\time\pp\hol\OpenPP.vbs"
	goto EX
		exit						
)
"D:\time\OpenPP.vbs" !test!
goto EX
)
exit
:EX
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
	IF !ERRORLEVEL!==0 (
		"D:\time\ex\CloseEX.vbs"
		timeout -t 10
		taskkill /IM EXCEL.EXE
		goto EX
	) 
exit