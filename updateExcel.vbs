Dim oApp
Dim oPres
Dim oSlide

Set objExcel = Object("Excel.Application")
'Set objWorkbook = objExcel.Workbooks("D:\Time\ex\time2.xlsm")

'objExcel.Application.Visible = True
'objExcel.Workbooks.Add
'objExcel.Cells(1, 1).Value = "Test value"

objExcel.Application.Run "D:\Time\ex\time2.xlsm!AutoEvent" 