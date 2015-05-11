Attribute VB_Name = "Services_Calibration"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub addCalibrationRecord()
Dim StrSQL As String

If IsNull(Form_Main.Weight1_5.Value) = True Or IsNull(Form_Main.Weight2_5) = True Or IsNull(Form_Main.Weight3_5.Value) = True Or IsNull(Form_Main.Weight4_5.Value) = True Or IsNull(Form_Main.Weight5_5.Value) = True Or IsNull(Form_Main.Weight1_20.Value) = True Or IsNull(Form_Main.Weight2_20.Value) = True Or IsNull(Form_Main.Weight3_20.Value) = True Or IsNull(Form_Main.Weight4_20.Value) = True Or IsNull(Form_Main.Weight5_20.Value) = True Then
    MsgBox "Process not completed!"
    Exit Sub
Else
    If IsNull(Form_Main.Name_calib) = True Then
        MsgBox "Please Insert your name to validate the calibration !"
    Else
        DoCmd.SetWarnings False
        StrSQL = " INSERT INTO [Calibration]([Date_calib],[Name_calib],[Workstation],[USER],[Weight1_5],[Weight2_5],[Weight3_5],[Weight4_5],[Weight5_5],[Weight1_20],[Weight2_20],[Weight3_20],[Weight4_20],[Weight5_20]) VALUES " _
                   & "('" & Form_Main.Date_calib & "', '" & Form_Main.Name_calib & "', '" & getWorkstation & "', '" & getUserName & "', '" & Form_Main.Weight1_5.Value & "', '" & Form_Main.Weight2_5.Value & "', '" & Form_Main.Weight3_5.Value & "', '" & Form_Main.Weight4_5.Value & "', '" & Form_Main.Weight5_5.Value & "', '" & Form_Main.Weight1_20.Value & "', '" & Form_Main.Weight2_20.Value & "', '" & Form_Main.Weight3_20.Value & "', '" & Form_Main.Weight4_20.Value & "', '" & Form_Main.Weight5_20.Value & "');"
        DoCmd.RunSQL StrSQL
        DoCmd.SetWarnings True
                    
        cleanCalibration
            
        Form_Main.Calibration_subform.Form.Requery
    End If
End If
End Sub

Public Sub cleanCalibration()
     Form_Main.Box195.BackColor = RGB(63, 63, 63)
     Form_Main.Box193.BackColor = RGB(237, 28, 36)
     Form_Main.Box196.BackColor = RGB(63, 63, 63)
     Form_Main.Box197.BackColor = RGB(63, 63, 63)
     Form_Main.Box194.BackColor = RGB(63, 63, 63)
     
     Form_Main.Weight1_5.Value = Null
     Form_Main.Weight2_5.Value = Null
     Form_Main.Weight3_5.Value = Null
     Form_Main.Weight4_5.Value = Null
     Form_Main.Weight5_5.Value = Null
     Form_Main.Weight1_20.Value = Null
     Form_Main.Weight2_20.Value = Null
     Form_Main.Weight3_20.Value = Null
     Form_Main.Weight4_20.Value = Null
     Form_Main.Weight5_20.Value = Null
     
     Form_Main.Name_calib.Value = Null
        
     Form_Main.poid.Caption = "5"
     Form_Main.calib_state.Value = "5_centre"
End Sub

Public Sub calibrationInput()
Dim StrSQL As String
Dim calib_state As String

On Error GoTo Routine_Error

If IsNull(Form_CalibrationInput.weightMeasuredData) = False And Form_CalibrationInput.weightMeasuredData > 0 Then

calib_state = Form_Main.calib_state.Value

'save weight in appropriate case
'update layout
'clean unbound

Select Case calib_state
    Case "5_centre"
        Form_Main.Weight1_5.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box193.BackColor = RGB(63, 63, 63)
        Form_Main.Box197.BackColor = RGB(237, 28, 36)
        calib_state = "5_up_left"
    Case "5_up_left"
        Form_Main.Weight2_5.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box197.BackColor = RGB(63, 63, 63)
        Form_Main.Box194.BackColor = RGB(237, 28, 36)
        calib_state = "5_up_right"
    Case "5_up_right"
        Form_Main.Weight3_5.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box194.BackColor = RGB(63, 63, 63)
        Form_Main.Box196.BackColor = RGB(237, 28, 36)
        calib_state = "5_down_left"
    Case "5_down_left"
        Form_Main.Weight4_5.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box196.BackColor = RGB(63, 63, 63)
        Form_Main.Box195.BackColor = RGB(237, 28, 36)
        calib_state = "5_down_right"
    Case "5_down_right"
        Form_Main.Weight5_5.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box195.BackColor = RGB(63, 63, 63)
        Form_Main.Box193.BackColor = RGB(237, 28, 36)
        Form_Main.poid.Caption = "20"
        calib_state = "20_centre"
        
    Case "20_centre"
        Form_Main.Weight1_20.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box193.BackColor = RGB(63, 63, 63)
        Form_Main.Box197.BackColor = RGB(237, 28, 36)
        calib_state = "20_up_left"
    Case "20_up_left"
        Form_Main.Weight2_20.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box197.BackColor = RGB(63, 63, 63)
        Form_Main.Box194.BackColor = RGB(237, 28, 36)
        calib_state = "20_up_right"
    Case "20_up_right"
        Form_Main.Weight3_20.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box194.BackColor = RGB(63, 63, 63)
        Form_Main.Box196.BackColor = RGB(237, 28, 36)
        calib_state = "20_down_left"
    Case "20_down_left"
        Form_Main.Weight4_20.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box196.BackColor = RGB(63, 63, 63)
        Form_Main.Box195.BackColor = RGB(237, 28, 36)
        calib_state = "20_down_right"
    Case "20_down_right"
        Form_Main.Weight5_20.Value = Form_CalibrationInput.weightMeasuredData
        Form_Main.Box195.BackColor = RGB(63, 63, 63)
        Form_Main.Box193.BackColor = RGB(237, 28, 36)
        Form_Main.poid.Caption = "5"
        calib_state = "5_centre"
        
     Case Else
        Debug.Print ("Calib_record_error")
End Select
Form_Main.calib_state.Value = calib_state

Call closeSerialPorts
DoCmd.Close acForm, "CalibrationInput"
Form_Main.next_btn.SetFocus
End
End If

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
 End
End Sub
