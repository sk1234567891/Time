REM check if saturday

date /t |findstr Sat
	IF %errorlevel%==0 (
						REM IF saturday this will run
							"D:\time\pp\Shabat\OpenPP.vbs"
						exit
						)
REM IF not saturday then this will run
	"D:\time\pp\hol\OpenPP.vbs"
exit						