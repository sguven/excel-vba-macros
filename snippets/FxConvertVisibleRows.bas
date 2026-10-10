Selection.SpecialCells(xlCellTypeVisible).Select



                              
    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
      Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-1]*(VLOOKUP(RC[-2],kur.xls!C1:C3,3,0))/kur.xls!R2C3"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lastRow)
