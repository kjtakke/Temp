Const WDdata as variant
Public XLWD as String

Sub InitiateDataArray
redim WDdata(0 to 999, 0 to 999,0 to 30,0 to 300)

'Metadata    0.0.0.0-#
'Page data   P.0.0.0-#
'Row data    P.R.0.0-#
'Column data P.R.C.0-#

Call GetWebPageData

Dim MTD as varient
Dim PGS,  RWS,  CLS as varient
Dim PGSD, RWSD, CLSD as varient


MTD = split(XLWD, "MM")
for M = 1 to ubound(MTD)
	WDdata(0,0,0,M) = MTD(M,1)
Next M

PGS = split(XLWD, "PP")
for P = 2 to ubound(Pages)
	PGSD = split(XLWD, "PD")
	for PD = 1 to ubound(PGSD)
		WDdata(P-1,0,0,PD) = PGSD(PD,1)
	next PD

	RWS = split(PGS(P,1))
	for R = 2 to ubound(Rows, "RR")
		RWSD = split(RWS(RD,1))
		for RD = 1 to ubound(RWSD)
			WDdata(P-1,R-1,0,RD) = RWSD(RD,1)
		next RD

		CLS = split(RWS(R,1))
		for C = 2 to ubound(CLS, "CC")
			CLSD = split(CLS(CD,1))
			for CD = 1 to ubound(CLSD)
				WDdata(P-1,R-1,C-1,CD) = CLSD(CD,1)
			next CD

		next C
	next R
next P
End Sub




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
#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long


    Public Declare PtrSafe Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Public Declare PtrSafe Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long
#Else
    Public Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long


    Public Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long


    Public Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" _
               (ByVal hwnd As Long) As Long
#End If
Sub HideBar(frm As Object)

Dim Style As Long, Menu As Long, hWndForm As Long
hWndForm = FindWindow("ThunderDFrame", frm.Caption)
Style = GetWindowLong(hWndForm, &HFFF0)
Style = Style And Not &HC00000
SetWindowLong hWndForm, &HFFF0, Style
DrawMenuBar hWndForm

End Sub
