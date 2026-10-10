r = trc.Range("B1").End(xlDown).Row + 1
trc.Range("B" & r).Value = trs.Range("E88").Value / usd
trc.Range("D" & r).Value = trc.Range("C" & r).Value / fxp.Range("AJ2").Value

