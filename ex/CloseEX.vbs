'Excel
On Error Resume Next

Set xl = GetObject(, "Excel.Application")
If Err Then
If Err.Number = 429 Then

WScript.Quit 0
Else

CreateObject("WScript.Shell").LogEvent 1, Err.Description & _
  " (0x" & Hex(Err.Number) & ")"
WScript.Quit 1
End If
End If
On Error Goto 0

xl.DisplayAlerts = False  

For Each wb In xl.Workbooks
wb.Save
wb.Close False
Next

xl.Quit
Set xl = Nothing