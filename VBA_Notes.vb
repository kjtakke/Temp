'https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
'

Sub OpenEdge()

ActiveWorkbook.FollowHyperlink Address:="microsoft-edge:https://www.google.com"

End Sub


'Name macro
Sub ListTables()

'Declare variables and data types
Dim tbl As ListObject
Dim WS As Worksheet
Dim i As Single, j As Single

'Insert new worksheet and save to object WS
Set WS = Sheets.Add

'Save 1 to variable i
i = 1

'Go through each worksheet in the worksheets object collection
For Each WS In Worksheets

    'Go through all Excel defined Tables located in the current WS worksheet object
    For Each tbl In WS.ListObjects
        
        'Save Excel defined Table name to cell in column A
        Range("A1").Cells(i, 1).Value = tbl.Name

        'Iterate through columns in Excel defined Table
        For j = 1 To tbl.Range.Columns.Count

            'Save header name to cell next to table name
            Range("A1").Cells(i, j + 1).Value = tbl.Range.Cells(1, j)

        'Continue with next column
        Next j

        'Add 1 to variable i
        i = i + 1
    
    'Continue with next Excel defined Table
    Next tbl

'Continue with next worksheet
Next WS

'Exit macro
End Sub

    TextBox1.MultiLine = True
    TextBox1.Value = [A1] & vbCr & [B1] & vbCr & [C1]

Replace(myText, vbCR, vbCRLF)

Sub passValuesToCell()
    Dim lines: lines = Split(UserForm1.TextBox25.value, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Sheet1.Range("J" & i + 1).Resize(, 3).value = Split(lines(i), vbTab)
    Next
End Sub

'Public Arrays KB
Public WDdata as variant            'Website Data Aarray
Public ColumnText as Variant        'Column Text Field
Public JavaScriptText as Variant    'JavaScript Field

'Deliminators
Const MD as String = "C:2s3ZpnC8A,S{T/H)SWZ'24\mmuv3Egb%M/QDA86AUer`zn=Z'u@8;tTry{gqYa5VK`.(y9LvR~&PTs\=RQW2<}A@s:#Lr>V(W2;-s4~$Wbq9~NT'},Q.bm*Rj'7Nve" 'Metadata
Const PP as String = "^cv8CR(U<3wbvh2>*ee.bK'6b:ZQqwj@s#?EQLhU:U>4Q:^[pALeg,/a+/]R$ZuG48_rTuC9)kQyKUZUe:#jv_.DK$3fm}g%*]~/,`A&$V;5;[yAz$BPw}TV`yXqB~G%" 'Page
Const PD as String = "C3`:j~52,`/Bt:b:y~y[^PRtznp8^XE-vSA:93=#LjLR>M~8%%$jB<x<G;5)*cB4sPFV9#}/Rd5E8^)<@NazNjEX8S~ND&Qk/Mt_n&3?Y5Dbxx[GNG#En,GZ&k-3RhD:" 'Page Data 
Const RR as String = "2@Nf<S>GH3VQEvZY+GSw:*-@(?%DV_h{#6AZp'6{DL`~w.cM<U$;8e'BqhyCpSZ2WQ'%}]N+6]xf`pT@,_b@a-g2[]*Hh!8}U4ngnYFVWgyV$y?::]D&bBw[fWD}Y~GF" 'Row
Const RD as String = "ED!>g9b)u&;4>?~>f'#3_x=S=Pjmstm:CZ/r'tY4A[fMk5~P%C>*77*)<u^9'sUXGWhKpZ9RtJ{%{zrABU4~Mrmh4MuS,,pGsSDEv4)[~F$M6PbCUEdA9gGgP'tbQzPn" 'Row Data
Const CC as String = "`YE%%8y8zGX_d7<y*FDSG3!h.\KF2qQf%A#z8[v\@ML~bU#ehM<U+aV3t,7YdwgU>ydR_E^V>4xzGXfP;c3j.a45QFJwRxv/pD:=5QEK~4@Q7m5]KkaD5;!#q_T6t$&'" 'Column
Const CD as String = "7,)zbZ,>-_*/*6yeLSX&~@@y,k6$HksXyX~ex}#g&(AL\yDY(kcn9!`xQg9$[GyVq;(a/vC$4T=^+jB?yL3M8m'u]76F)v/XaEV'#>K?f5g=5]7>^h!y4:%c{_SW*fky" 'Column Data
Const TD as String = "auY5xUf454Ajc6~~}.CS.<DZU7bB?Ee.+;YZ5$J?N9!68.~fgrquYj]{A,5Rfe$(;=caBe*\g!$%b4REtwkn6w]]cT>N[T([VE_J?%}$DNak`w)@:58zse[<4M#d.Zp6" 'Text Column Data
Const JD as String = "Hn:w_N<mSpmm,w~_DrC,~}:6$yneD:9+KhA>,nr3X+w-:jVQYCpND=]4?-,g[pA)wcN''zffZW(U=?&uXGj&~V%8N^5ryBN`@+MsY!<;x`;r6dd#y*'5@:x2{u_w}LH]" 'JavaScript Data

'Array Dimentions
Const AD1 as Integer = 100  'WDdata Dimention 1
Const AD2 as Integer = 100  'WDdata Dimention 2
Const AD3 as Integer = 30   'WDdata Dimention 3
Const AD4 as Integer = 300  'WDdata Dimention 4

'Data Fields
Const DM as integer = 300   'Metadata Fields
Const DP as integer = 300   'Page Fields
Const DR as integer = 300   'Row Fields
Const DC as integer = 300   'Column Fields
Public XLWD as String

Sub ConcatinateWDDataArray
    'P = Page    R = Row      C = Column

    'Metadata    0.0.0.0-##
    'Page data   P.0.0.0-##   0 = 0(exclude) or 1(include)
    'Row data    P.R.0.0-##   0 = 0(exclude) or 1(include)
    'Column data P.R.C.0-###  0 = 0(exclude) or 1(include)

    Dim MTD as varient
    Dim AllData as String
    'WDdata(0 to 100, 0 to 100, 0 to 30, 0 to 300)

    AllData = ""

    'Metadata
    For M=0 to DM
        AllData = AllData & WDData(0,0,0,M) & MM
    Next M
    AllData = AllData & PP
    'Pages
    For P = 0 to 100

        'Page Data
        For PA = 0 to DP
            AllData = AllData & WDData(P,0,0,PA) & PD
        Next PA
        
        'Rows
        For R = 0 to 100 

            'Row Data
            For RA = 0 to DR
                AllData = AllData & WDData(P,R,0,RA) & RD
            Next RA

            'Columns
            For C = 1 to 30
                'Column Data
                For CA = 0 to DC 
                    AllData = AllData & WDData(P,R,C,CA) & CD
                Next CA

                'Concatinate Columns
                AllData = AllData & CC
            Next C

            'Concatinate Rows
            AllData = AllData & RR
        Next R

        'Concatinate Pages
        AllData = AllData & PP
    Next P

    'Write to XLWD File
    XLWD = AllData
    Call WriteXLWD(XLWD)
End Sub


Sub SplitWDDataArray
    redim WDdata(0 to AD1, 0 to AD2, 0 to AD3, 0 to AD4)

    'P = Page    R = Row      C = Column

    'Metadata    0.0.0.0-##
    'Page data   P.0.0.0-##   0 = 0(exclude) or 1(include)
    'Row data    P.R.0.0-##   0 = 0(exclude) or 1(include)
    'Column data P.R.C.0-###  0 = 0(exclude) or 1(include)

    Call GetWebPageData

    Dim MTD as varient
    Dim PGS,  RWS,  CLS as varient
    Dim PGSD, RWSD, CLSD as varient

    'Metadata
    MTD = split(XLWD, MM)
    for M = 1 to ubound(MTD)-1
        WDdata(0,0,0,M-1) = MTD(M,1)
    Next M

    'Pages
    PGS = split(XLWD, PP)
    for P = 2 to ubound(PGS)

        'Page Data
        PGSD = split(PGS(P,1), PD)
        for PA = 1 to ubound(PGSD) - 1
            WDdata(P-1,0,0,PA-1) = PGSD(PA,1)
        next PA
            
        'Rows
        RWS = split(PGS(P,1), RR)
        for R = 2 to ubound(RWS)

            'Row Data
            RWSD = split(RWS(R,1),RD)
            for RA = 1 to ubound(RWSD) - 1
                WDdata(P-1,R-1,0,RA-1) = RWSD(RA,1)
            next RA

            'Columns
            CLS = split(RWS(R,1),CC)
            for C = 2 to ubound(CLS)

                'Column Data
                CLSD = split(CLS(C,1),CD)
                for CA = 1 to ubound(CLSD) - 1
                    WDdata(P-1,R-1,C-1,CA-1) = CLSD(CA,1)
                next CA
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
