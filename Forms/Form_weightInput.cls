VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_weightInput"
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


Private Sub exitWeightInput_Click()


On Error GoTo Routine_Error

Call closeSerialPorts
DoCmd.Close acForm, "weightInput"
Form_Main.SSCCNumber.SetFocus

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End

End Sub


Private Sub Command12_Click()
Call closeSerialPorts
DoCmd.Close acForm, "weightInput"
Form_Main.SSCCNumber.SetFocus
End
End Sub

Private Sub weightMeasuredData_AfterUpdate()

On Error GoTo Routine_Error

If IsNull(Me.weightMeasuredData) = False And Me.weightMeasuredData > 0 Then


Form_Main.measuredWeight = Me.weightMeasuredData

Call closeSerialPorts
DoCmd.Close acForm, "weightInput"
Form_Main.SSCCNumber.SetFocus

checkingSequence "manual"
End
End If

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
    End
    
End Sub
