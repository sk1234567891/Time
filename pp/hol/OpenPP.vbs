Dim oApp
Dim oPres
Dim oSlide

Set oApp = CreateObject("Powerpoint.Application")
oApp.visible = true
Set oPres = oApp.Presentations.Open("D:\time\pp\hol\hol.pptm")
oApp.Run "hol.pptm!updatelink"
oPres.Save
oPres.SlideShowSettings.Run
