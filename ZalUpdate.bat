cd /D D:\time
	REM Start the time file for updates
start EXCEL.EXE ex\zal.xlsx
timeout -t 50
rem pause
	REM save the time file
"D:\time\ex\CloseEX.vbs"
	REM Exit the time file
taskkill /IM EXCEL.EXE