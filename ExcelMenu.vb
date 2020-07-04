'const numeric = [
    '['input1','input2'],
    '['input3','input4']
'];

var numeric = [
    ['input1','input2'],
    ['input3','input4']
];
'numeric[0][0] == 'input1';
'numeric[0][1] == 'input2';
'numeric[1][0] == 'input3';
'numeric[1][1] == 'input4';

function JSA(data as varient)
dim D as long 
dim tempAry as varient
tempAry = data
D = dimentions(tempAry)
tempAry = empty

dim V as string
V= "V"
dim H as string
V= "H"

data = 

JS = "const " & varNm & "[" & vbnewline


for i  1 to ubound(data)
JS = JS & "["
for j = 1 to default
if isnumeric(data(i,j)) = true then JS = JS & data(i,j) & "," & vbnewline
else
JS = JS & "'" & data(i,j) & "'," & vbnewline
end if
next j
JS = JS & "],"
next i
JS = JS & "];"
end function










https://www.onlinegdb.com/online_vb_compiler

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("&MyFunction").Delete
      On Error GoTo 0
   End With
End Sub

Private Sub Workbook_Open()
   Dim objPopUp As CommandBarPopup
   Dim objBtn As CommandBarButton
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("MyFunction").Delete
      On Error GoTo 0
      Set objPopUp = .Controls.Add( _
         Type:=msoControlPopup, _
         before:=.Controls.Count, _
         temporary:=True)
   End With
   objPopUp.Caption = "&MyFunction"
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Formula Entry"
      .OnAction = "Cbm_Active_Formula"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Value Entry"
      .OnAction = "Cbm_Active_Value"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Formula Selection"
      .OnAction = "Cbm_Formula_Select"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Value Selection"
      .OnAction = "Cbm_Value_Select"
      .Style = msoButtonCaption
   End With
End Sub
