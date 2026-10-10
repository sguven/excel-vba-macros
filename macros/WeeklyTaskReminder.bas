
Sub Macrodeneme1()

If Weekday(Date) = 6 Then
    Dim Msg, Style
    Msg = "WEEKLY YAPILDI MI"
    Style = vbYesNo + vbDefaultButton1 + vbCritical
    Response = MsgBox(Msg, Style)
    If Response = vbYes Then
        
        'YES DERSEN BOYUYOR HÜCREYİ MESELA
        
        Range("E5").Interior.Color = 49407
        Range("E5") = "BUGÜN PERŞEMBE VE GÖREV TAMAM"
        
    
    Else
        
        'BURDA BİŞİ YAPMAZ
        Range("E5") = "BUGÜN PERŞEMBE AMA UNUTTUN KRAL"

        
            
    
    End If
    
    
Else
    Range("E5") = "BUGÜN PERŞEMBE DEĞİLMİŞ RAAD"
    
    
End If

End Sub
