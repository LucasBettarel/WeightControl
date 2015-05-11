VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_manual check pop up"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Antoine NICOLE
'Stephen HOUSSAYE
'Lucas BETTAREL

Option Compare Database


Private Sub manualExplanationList_AfterUpdate()

If Me.manualExplanationList.Value <> "" Then
Me.manualExplanationOthers = ""
Else
Me.manualExplanationList.Value = ""
End If

End Sub

Private Sub manualExplanationOthers_AfterUpdate()

If Me.manualExplanationOthers <> "" Then
Me.manualExplanationList.Value = ""
Else
Me.manualExplanationOthers = ""
End If

End Sub

Private Sub manualOK_click()
Dim theoWeight As Double
Dim actualWeight As Double
theoWeight = 0
actualWeight = 0

Dim Comment As String

Comment = ""

If Me.manualExplanationList.Value <> "" Then Comment = Me.manualExplanationList.Value
If Me.manualExplanationOthers <> "" Then Comment = Me.manualExplanationOthers

If IsNull(Comment) = True Or Comment = "" Then
MsgBox "Please enter a raison for manual check"
Exit Sub
End If

recordCheckResult "Manual", "Pass", theoWeight, actualWeight, Comment, 0, ""

DoCmd.Close acForm, "manual check pop up"

cleaningWC

  If Form_Main.weightStationCheck = -1 Then

Form_Main.SSCCNumber.SetFocus
initiateAndon
  
  Else
Form_Main.Material_Control.SetFocus
initiateAndon
  End If

If Form_Main.EANautomaticprinting = -1 Then

Dim sql2 As String
Dim workRS2 As Recordset

sql2 = "Select * " & _
        "FROM [printers] " & _
        "WHERE [workstation] = '" & getWorkstation & "';"

Set workRS2 = CurrentDb.OpenRecordset(sql2)

If workRS2.RecordCount = 0 Then Exit Sub

selectedPrinter = workRS2![Printer]
On Error GoTo Routine_Error
Set Application.Printer = Application.Printers(selectedPrinter)

selectLabel = "EAN13CheckLabel"
DoCmd.SelectObject acReport, Trim(selectLabel), True
DoCmd.PrintOut , , , , 1
End If

Routine_Exit:
    Exit Sub

Routine_Error:
    MsgBox "The printer has not been found"
    Resume Routine_Exit
End Sub
