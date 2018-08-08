Dim oApp
Dim oPres
Dim oSlide

Set oApp = CreateObject("Powerpoint.Application")
oApp.visible = true
Set oPres = oApp.Presentations.Open("D:\time\pp\Shabat\show2.pptm")
oApp.Run "show2.pptm!updatelink"
oPres.Save
oPres.SlideShowSettings.Run
