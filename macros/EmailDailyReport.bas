Sub email()

Dim x As Variant
Dim day, month, year As String
x = ThisWorkbook.Sheets("Sheet1").Range("J2").Value

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


Workbooks.Open Filename:="J:\adress\file_" & day & month & year & ".xlsx"

Workbooks("file_" & day & month & year & ".xlsx").Activate
Sheets("Tablo").Select
Range("B3:E11").Select
Selection.Copy

Dim outlookApp As Outlook.Application
Dim myMail As Outlook.MailItem
Dim Source_File As String

Set outlookApp = New Outlook.Application
Set myMail = outlookApp.CreateItem(olMailItem)

value = Workbooks("file_" & day & month & year & ".xlsx").Worksheets("Sheet1").Range("E31")
word = Workbooks("file_" & day & month & year & ".xlsx").Worksheets("Sheet1").Range("E32")

myMail.To = "blabla@email.com"
myMail.CC = "blabla@email.com"

myMail.Subject = "my email " & x

'use HTML here
myMail.HTMLBody = "<p style='font-family:calibri' >Dear all,</p><p style='font-family:calibri'>The total value " & word & " by " & value & " USD. </p><p style='font-family:calibri'> Best Regards. </p>"

'to add body "<b>This is bold</b><br> and this isn't"

'myMail.Attachments.Add "J:\adress\file_" & day & month & year & ".xlsx"

myMail.Display True 'to show the e-mail before sending
'myMail.Send 'to send directly

End Sub
