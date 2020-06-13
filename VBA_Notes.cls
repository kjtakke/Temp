Function AryFromCell(WS as string, index as string, optional values as string) As Varient

dim tmpAry, rngAry as variant
dim H as long
dim rng as range
dim cAdd1, cAdd2 as string

if values = "" then

AryFromCell = worksheets(WS).range(index, worksheets(WS).range(index).end(xldown)).value

else

set rng = range(values)

cAdd1 = rng.Row & "," & rng.Column

rngAry = split(cAdd, ",")

tmpAry = worksheets(WS).range(index, worksheets(WS).range(index).end(xldown)).value

H = rngAry(0) + ubound(tmpAry)

cAdd2 = H & ", " & rngAry(1)

AryFromCell = worksheets(WS).range(cells(cAdd1), cells(cAdd2)).value

end if
End Function
