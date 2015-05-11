VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
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

Private Sub aboutbtn_Click()
    DoCmd.OpenForm "About", acNormal
End Sub

Private Sub add_calibration_record_Click()
    addCalibrationRecord
End Sub
Private Sub checkCalculatedWeight_Click()

If checkCalculatedWeight.Value = 0 Then
Me.checkSAPWeight.Value = -1
Me.checkCalculatedWeight.Value = 0
Else
Me.checkSAPWeight.Value = 0
Me.checkCalculatedWeight.Value = -1
End If
End Sub

Private Sub checkSAPWeight_Click()

If Me.checkSAPWeight.Value = -1 Then
Me.checkSAPWeight.Value = -1
Me.checkCalculatedWeight.Value = 0
Else
Me.checkSAPWeight.Value = 0
Me.checkCalculatedWeight.Value = -1
End If
End Sub

Private Sub Command1038_Click()
result = MsgBox("Are you sure that you want to quit Weight Control ?", vbYesNo)
If result = vbYes Then
    DoCmd.Quit
Else
    Exit Sub
End If
End Sub

Private Sub Command110_Click()

DoCmd.OpenTable "boxWeight", acViewNormal, acEdit

End Sub

Private Sub Command114_Click()

'recordMasterdata issue
recordMDissue
End Sub

Private Sub Command115_Click()
DoCmd.OpenForm "Individual Weight Update", acNormal
End Sub

Private Sub Command12_Click()
uploadMaterialMaster
End Sub

Private Sub Command14_Click()
DoCmd.OpenTable "Material Master", acViewNormal, acEdit
End Sub

Private Sub Command15_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, acSpreadsheetTypeExcel12, "Material Master", "Microsoft Excel Workbook (*.xlsx)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub Command167_Click()
DoCmd.OpenForm "loginForm", acNormal
Form_loginForm.Visible = True
End Sub

Private Sub Command168_Click()
DoCmd.OpenForm "Feedback", acNormal
Form_Feedback.Move 7000, 5000
End Sub

Private Sub Command27_Click()
initiateEANcheck
End Sub

Private Sub Command32_Click()
Call closeSerialPorts
End
End Sub

Private Sub Command33_Click()
DoCmd.OpenTable "printers", acViewNormal, acEdit
End Sub

Private Sub Command35_Click()
DoCmd.OpenTable "parameters", acViewNormal, acEdit
End Sub

Private Sub Command37_Click()
DoCmd.OpenTable "logonData", acViewNormal, acEdit
End Sub

Private Sub Command41_Click()
uploadQCcheckLog
End Sub

Private Sub Command42_Click()
DoCmd.OpenTable "QCcheckLog", acViewNormal, acEdit
End Sub

Private Sub Command43_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, "QCcheckLog", "Microsoft Excel 97-2003 (*.xls)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub Command56_Click()
DoCmd.OpenTable "pickerList", acViewNormal, acEdit
End Sub

Private Sub Command58_Click()
uploadPickers
End Sub

Private Sub Command59_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, "pickerList", "Microsoft Excel 97-2003 (*.xls)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub Command60_Click()
DoCmd.OpenTable "sensitiveMaterial", acViewNormal, acEdit
End Sub

Private Sub Command62_Click()
uploadSensitive
End Sub

Private Sub Command63_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, "sensitiveMaterial", "Microsoft Excel 97-2003 (*.xls)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub Command76_Click()
DoCmd.OpenTable "weightControlSpecifications", acViewNormal, acEdit
End Sub

Private Sub Command82_Click()
'TODO #6
recordPickingIssues
End Sub

Private Sub Command83_Click()
uploadPickingIssues
End Sub

Private Sub Command84_Click()
DoCmd.OpenTable "pickingIssuesReport", acViewNormal, acEdit
End Sub

Private Sub Command85_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, "pickingIssuesReport", "Microsoft Excel 97-2003 (*.xls)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub Command88_Click()
updateKPI
End Sub
Private Sub Command94_Click()
DoCmd.OpenTable "KPI", acViewNormal, acEdit
End Sub

Private Sub Command95_Click()
uploadLocalMaterialMaster
End Sub

Private Sub Command96_Click()
DoCmd.OpenTable "Local Material Master", acViewNormal, acEdit
End Sub

Private Sub Command97_Click()
On Error GoTo cmdoutputTable_Error
DoCmd.OutputTo acOutputTable, "Local Material Master", "Microsoft Excel 97-2003 (*.xls)"
cmdoutputTable_Error:
Exit Sub
End Sub

Private Sub EANautomaticprinting_AfterUpdate()
If Me.EANautomaticprinting = -1 Then
      Me.weightAutomaticPrinting = -1
Else
      Me.weightAutomaticPrinting = 0
End If
End Sub

Private Sub GoCheck_Click()
ValidateWCTest
End Sub

Private Sub Help_btn_Click()
DoCmd.OpenForm "Help", acNormal
End Sub

Private Sub next_btn_Click()

On Error GoTo Routine_Error

DoCmd.OpenForm "CalibrationInput", acNormal
Form_CalibrationInput.Move 6000, 4000

 Call closeSerialPorts
 Call initSerialPorts
 Call ProcessDataFlow
 Call closeSerialPorts

DoCmd.Close acForm, "CalibrationInput"

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End
End Sub

Private Sub weightAutomaticPrinting_AfterUpdate()
    If Me.weightAutomaticPrinting = -1 Then
            Me.EANautomaticprinting = -1
    Else
            Me.EANautomaticprinting = 0
    End If
End Sub

Private Sub Form_Load()
  DoCmd.ShowToolbar "Ribbon", acToolbarNo
        DoCmd.ShowToolbar "Menu Bar", acToolbarNo
        CurrentDb.Properties("StartUpShowDBWindow") = False
updateKPI
End Sub

Private Sub KPI_AfterUpdate()
ResGraphCalcul
End Sub

Private Sub Loock_AfterUpdate()
LockApp
End Sub

Private Sub MonthLenght_AfterUpdate()
ResGraphCalcul
End Sub

Private Sub ResTimeLenght_AfterUpdate()
ResGraphCalcul
End Sub

Private Sub ResTimeUnit_AfterUpdate()
ResGraphCalcul
End Sub

Private Sub SSCCNumber_AfterUpdate()
cleaningWC

If Len(Me.SSCCNumber) <> 18 And Len(Me.SSCCNumber) <> 20 Then
 MsgBox "The SSCC number entered is incorrect"
 Exit Sub
End If
 
If Len(Me.SSCCNumber) = 20 Then
 temp = Right(Me.SSCCNumber, 18)
 Me.SSCCNumber = temp
End If
 
Me.SSCCnumber2 = Me.SSCCNumber

RFC_CALL_TRANSACTION_USING
 
If Me.pickerSESA = "" Then MsgBox "The picker SESA is unknown, please check"
 
useLocalData

Me.HandlingUnitsSubform.Requery

checkingSequence ("auto")

End Sub

Private Sub SSCCnumber2_AfterUpdate()
 
If Len(Me.SSCCnumber2) <> 18 And Len(Me.SSCCnumber2) <> 20 Then
 MsgBox "The SSCC number entered is incorrect"
 Exit Sub
End If
If Len(Me.SSCCnumber2) = 20 Then
 temp = Right(Me.SSCCnumber2, 18)
 Me.SSCCnumber2 = temp
End If
Me.SSCCNumber = Me.SSCCnumber2

RFC_CALL_TRANSACTION_USING
useLocalData
Me.HandlingUnitsSubform.Requery

updateWaitingList
updateHUWeight
initiateEANcheck
End Sub
