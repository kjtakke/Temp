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























Option Explicit
Private Declare Function GetSysColor Lib "user32.dll" ( _
ByVal nIndex As Long) As Long
Private Declare Function SetSysColors Lib "user32.dll" ( _
ByVal nChanges As Long, _
ByRef lpSysColor As Long, _
ByRef lpColorValues As Long) As Long
Const COLOR_ACTIVECAPTION As Long = 2
Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Const COLOR_CAPTIONTEXT As Long = 9
Dim myCOLOR_CAPTIONTEXT As Long, myCOLOR_ACTIVECAPTION As Long
Dim myCOLOR_GRADIENTACTIVECAPTION As Long
-----------------------------------------
Paste this in main userform module after creating two command buttons on the
test userform:
-----------------------------------------

Private Sub CommandButton1_Click()
Dim lngReturn As Long
myCOLOR_CAPTIONTEXT = GetSysColor(COLOR_CAPTIONTEXT)
myCOLOR_ACTIVECAPTION = GetSysColor(COLOR_ACTIVECAPTION)
myCOLOR_GRADIENTACTIVECAPTION = GetSysColor(COLOR_GRADIENTACTIVECAPTION)

lngReturn = SetSysColors(1, COLOR_CAPTIONTEXT, &HC0FFC0) 'Hex values for
the primary colors
lngReturn = SetSysColors(1, COLOR_ACTIVECAPTION, vbWhite)
lngReturn = SetSysColors(1, COLOR_GRADIENTACTIVECAPTION, vbRed)
End Sub

Private Sub CommandButton2_Click()
Dim lngReturn As Long

lngReturn = SetSysColors(1, COLOR_CAPTIONTEXT, myCOLOR_CAPTIONTEXT)
lngReturn = SetSysColors(1, COLOR_ACTIVECAPTION, myCOLOR_ACTIVECAPTION)
lngReturn = SetSysColors(1, COLOR_GRADIENTACTIVECAPTION,
myCOLOR_GRADIENTACTIVECAPTION)
End Sub
