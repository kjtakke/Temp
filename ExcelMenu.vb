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
