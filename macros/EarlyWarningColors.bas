Sub EWI()

Select Case Left(ewi_c, 1)

    Case "g"
    
    wb.Sheets(1).Range("E25").Interior.Color = 5296274
    
    Case "o"
    
    wb.Sheets(1).Range("E25").Interior.Color = 49407
    
    Case "r"
    
    wb.Sheets(1).Range("E25").Interior.Color = 255
    
End Select


Select Case Mid(ewi_c, 2, 1)

    Case "g"
    
    wb.Sheets(1).Range("K25").Interior.Color = 5296274
    
    Case "o"
    
    wb.Sheets(1).Range("K25").Interior.Color = 49407
    
    Case "r"
    
    wb.Sheets(1).Range("K25").Interior.Color = 255
    
End Select

Select Case Mid(ewi_c, 3, 1)

    Case "g"
    
    wb.Sheets(1).Range("Q25").Interior.Color = 5296274
    
    Case "o"
    
    wb.Sheets(1).Range("Q25").Interior.Color = 49407
    
    Case "r"
    
    wb.Sheets(1).Range("Q25").Interior.Color = 255
    
End Select

Select Case Mid(ewi_c, 4, 1)

    Case "g"
    
    wb.Sheets(1).Range("L5").Interior.Color = 5296274
    
    Case "o"
    
    wb.Sheets(1).Range("L5").Interior.Color = 49407
    
    Case "r"
    
    wb.Sheets(1).Range("L5").Interior.Color = 255
    
End Select
End Sub
