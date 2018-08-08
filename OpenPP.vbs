Dim oApp
Dim oPres
Dim oSlide

Set oApp = CreateObject("Powerpoint.Application")
oApp.visible = true
Set oPres = oApp.Presentations.Open("D:\time\pp\hol\" &WScript.Arguments(0) &".pptm")
oApp.Run (WScript.Arguments(0) &".pptm!updatelink")
oPres.Save
oPres.SlideShowSettings.Run
