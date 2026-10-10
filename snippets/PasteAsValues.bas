 Range("F65:X70").Select
 Selection.Copy
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
 Application.CutCopyMode = False
