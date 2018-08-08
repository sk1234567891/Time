Dim oApp
Dim oPres
Dim oSlide
' Dim oOpres



Set oApp = CreateObject("Powerpoint.Application")
oApp.visible = true

Set oPres = oApp.Presentations.Open("D:\time\pp\hol\" & WScript.Arguments(0), 2, True)

' Set oPres = oApp.Presentations.Open(objArgs)
' Set FileRun = 
oApp.Run(WScript.Arguments(0) &"!updatelink")
oPres.Save
oPres.SlideShowSettings.Run
