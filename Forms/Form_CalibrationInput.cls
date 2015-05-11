VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CalibrationInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database
Private Sub exitWeightInput_Click()

On Error GoTo Routine_Error


Call closeSerialPorts
DoCmd.Close acForm, "CalibrationInput"
End

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End
  
End Sub

Private Sub Command12_Click()
On Error GoTo Routine_Error


Call closeSerialPorts
DoCmd.Close acForm, "CalibrationInput"
End

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End

End Sub
Private Sub weightMeasuredData_AfterUpdate()
calibrationInput
End Sub
