On Error Resume Next
Set objPPT = GetObject(,"powerpoint.Application")
On Error Goto 0

If Not IsEmpty(objPPT) Then
For Each doc In objPPT.Presentations
doc.Save
doc.Close

Next

objPPT.Quit
End If