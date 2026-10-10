Sub Macro1()

Dim x As Variant
Dim day, month, year As String
x = InputBox("Enter Date:  dd/mm/yyyy", "Report Date")
day = Left(x, 2)
month = Mid(x, 4, 2)
year = Right(x, 2)

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

Workbooks.Open Filename:="J:\adress\" & "file" & day & month & year & ".xlsx" & ""

try = Workbooks("file" & day & month & year & ".xlsx").Worksheets("Tables").Range("F32").Value
usd = Workbooks("file" & day & month & year & ".xlsx").Worksheets("Tables").Range("D32").Value
eur = Workbooks("file" & day & month & year & ".xlsx").Worksheets("Tables").Range("H32").Value
atm = Workbooks("file" & day & month & year & ".xlsx").Worksheets("Tables").Range("S32").Value

Workbooks("file" & day & month & year & ".xlsx").Worksheets("Sheet1").Range("A1") = "bla bla bla " & x & ". 2Y TRY yields changed by " & try & " bps. 2Y USD yields changed by " & usd & " bps. 2Y EUR yields changed by " & eur & " bps. 3 month ATM volatility is " & atm & "."""

End Sub
