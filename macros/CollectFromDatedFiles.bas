Sub loop()

Dim i As Integer
Dim a As Integer
Dim day, month, year As String

'a defines that how many cells will be updated
a = ThisWorkbook.Sheets("Sheet1").Range("G8").Value

i = 1
Do While i < a

day = Left(Cells(i, 1), 2)
month = Mid(Cells(i, 1), 4, 2)
year = Right(Cells(i, 1), 2)

If month = 1 Then
month1 = "January"
End If
If month = 2 Then
month1 = "February"
End If
If month = 3 Then
month1 = "March"
End If
If month = 4 Then
month1 = "April"
End If
If month = 5 Then
month1 = "May"
End If
If month = 6 Then
month1 = "June"
End If
If month = 7 Then
month1 = "July"
End If
If month = 8 Then
month1 = "August"
End If
If month = 9 Then
month1 = "September"
End If
If month = 10 Then
month1 = "October"
End If
If month = 11 Then
month1 = "November"
End If
If month = 12 Then
month1 = "December"
End If

Dim year1 As String
year1 = Right(year, 2)

Workbooks.Open Filename:="J:\user\file\" & month1 & "-" & year1 & "\" & "20" & year & month & day & "\" & "20" & year & month & day & "-PVBP" & "\actual_file.xlsx"
  
Sheets("Sheet1").Select
Range("A1").Select
Selection.Copy
Windows("loop macro.xlsm").Activate
Cells(i, 2).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Windows("actual_file").Activate
ActiveWindow.Close

i = i + 1

Loop

End Sub
