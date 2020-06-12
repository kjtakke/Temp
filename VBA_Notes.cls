Function boxChart( _
	srtAry as Variant _
)as Variant

reDim boxChart(0 to 4)

boxChart(4) = srtAry(ubound(srtAry))

boxChart(0) = srtAry(0)

boxChart(1) = Round( _
	boxChart(4) / 4, 0)
	
boxChart(3) = boxChart(1) * 3

boxChart(2) = boxChart(1) * 2

End Function
'###########
Function mdAry_sdAry( _
	ary as Variant _
) as variant

reDim mdAry_sdAry( _
	0 to ubound(ary)-1)
	
For i = 1 to ubound(ary)
	mdAry_sdAry(i-1) = ary(i,1)
Next i

End Function
'###########
Function strRng_sdAry( _
	rng as string _
	ws as string
) as variant

dim ary as Variant
ary = worksheets(ws).range(rng).value

reDim strRng_sdAry( _
	0 to ubound(ary)-1)
	
For i = 1 to ubound(ary)
	strRng_sdAry(i-1) = ary(i,1)
Next i

End Function
'###########
Function bxMlti( _
	index as string, _
	values as string, _
	ws as string _
) as Varient

Dim iAry, vAryas variant
iAry = worksheets(ws).range(index).value

vAry = worksheets(ws).range(values).value

Dim uniqueAry as Variant
uniqueAry = unique(iAry, true)

reDim bxMltiy( _
1 to ubound(uniqueAry), 1 to 6)


End Funcrion
