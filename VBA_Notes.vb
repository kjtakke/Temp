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


function twoDP(agg as string, index as variant, values as variant) as variant

dim ary1, ary2, unique as variant
dim c as long
dim avg as variant
dim isSame as boolean 
redim unique(1 to ubound(index))

'Unique value set
unique(1) = index(1,1)
c = 1
for i = 2 to ubound(index)
	isSame = false
	for j = 1 to ubound(unique)
		if unique(j) = index(i,1) then 			isSame = true
	next j
	if isSame = false then
	c = c + 1
	unique(c) = index(i,1)
next i
redim preserve unique(1 to c, 1 to 2)

'Aggregation
for i = 1 to ubound(unique)
	redim avg(1 to 2)
	avg(1) = 0
	avg(2) = 0
	for j = 1 to ubound(values)
		if index(j,1) = unique(i,1) then
			avg(1) = avg(1) + 1
			avg(2) = avg(2) + values(j,1)
		end if
	next j
	select case true
		case agg = "sum"
			unique(i,2) = avg(2)
		case agg = "average"
			unique(i,2) = Round(avg(2)/avg(1),0)
		case agg = "count"
			unique(i,2) = avg(1)
	end select
next i

'return
twoDP = unique
end function



function TwoTxtDataPoints(ary as variant) as string

TwoTxtDataPoints = ""

For i = 1 to ubound(index)

	TwoTxtDataPoints = TwoTxtDataPoints & "{ label: '" & ary(i,1) & "' y: " & ary(i,2) & " }," & vbnewline

next i

end function


function TwoNoDataPoints(ary as variant) as string

TwoDataPoints = ""

For i = 1 to ubound(index)

	TwoNoDataPoints = TwoNoDataPoints & "{ label: " & ary(i,1) & " y: " & ary(i,2) & " }," & vbnewline

next i

end function



function BubbleDataPoints(ary as variant) as string

BubbleDataPoints = ""

For i = 1 to ubound(index)

	BubbleDataPoints = BubbleDataPoints & "{  name: ''" & ary(i,1) & "' x: " & ary(i,2) & " y: " & ary(i,3) & " z: " & ary(i,4) & " }," & vbnewline

next i

end function



function BoxDataPoints(ary as variant) as string

BoxDataPoints = ""

For i = 1 to ubound(index)

	BoxDataPoints = BoxDataPoints & "{ label: '" & ary(i,1) & "' y: " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & " }," & vbnewline

next i

end function

function RangePoints(ary as variant) as string

RangePoints = ""

For i = 1 to ubound(index)

	RangePoints = RangePoints & "{ label: '" & ary(i,1) & "' y: " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & ", " & ary(i,2) & " }," & vbnewline

next i

end function
