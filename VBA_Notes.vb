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
For 


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



Function TwoArraysToOneDataSet(ary1 as variant, ary2 as varient) as variant

dim ary as variant
redim ary(1 to ubound(ary1), 1 to 2)

for i = 1 to ubound(ary1)

ary(i,1) = ary1(i,1)
ary(i,2) = ary2(i,1)

next i

TwoArraysToOneDataSet = ary

end function


Function chart( _
		  index as string, _
		  values as string, _
		  chartType as string, _
optional kpi as string, _
optional dimentions as string, _
optional chartTitle as string, _
optional chartSubtitles as string, _
optional legend as string, _
optional toolTip as string, _
optional colors as string, _
optionap theme as string, _
optional backgroundColour as string, 
optional startAtZero as boolean, _
optional boundries as string, _
optional prefix as string, _
optional sufix as string, _
optional animation as string, _
optional 
) as variant

end function


function extractData(fileLocas string) as string

'Use txt file as storage
'File picker open and save
'Format .wd

'Deliminators
'Page: !~P~!
'Row: !~R~!
'Column: !~C~!
'Other: !~|~! !~||~! !~|||~!

'1.1.1 = Page data
'1.2(50).1 = Row data
'1.2(50)0.2(20) = Column data

'2.1.1 = Page data
'2.2(50).1 = Row data
'2.2(50)0.2(20) = Column data

'3.1.1 = Page data
'3.2(50).1 = Row data
'3.2(50)0.2(20) = Column data


'GET FILE DATA
dim getData as variant




'SPLIT PAGES
dim pData as variant
pData = split(getData, "!~P~!")


'SPLIT ROWS
dim rData, tmpAry1, tmpAry2 as variant
dim tmpStr as String
redim rData(1 to 100, 1 to 100)

for i = 1 to ubound(pData)

r = 1
tmpStr = pData(i,1)
tmpAry1 = split(tmpStr, "!~R~!")
rData(i,1) = tmpAry(1,1)

for j = 2 to ubound(tmpAry)+1

rData(i,j+1) = tmpAry1(j,1)

next j

next i

'SPLIT COLUMNS
dim cData, cStr as varient
redim cData(1 to 100, 1 to 100, 1 to 30)
dim r, c as integer

for i = 1 to ubound(pData)
cData(i,1,1) = rData(i,1)

'r = deturmin dimentions of each row
for j = 2 to r
cData(i,j,1) = rData(j,1)
cstr = split(rData(i,j),"!~C~!")

for k = 3 to ubound(cStr)+3
cData(i,j,k) = cStr(k-3)

next k
next j
next i

extractData = cData
end Function

