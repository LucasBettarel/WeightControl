Attribute VB_Name = "Views_Main_WC"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub cleaningWC()

Form_Main.loadingWeight = 0
Form_Main.tareWeight = 0
Form_Main.totalWeight = 0
Form_Main.measuredWeight = 0

Form_Main.packagingType = ""
Form_Main.loadingWeight2 = 0
Form_Main.tareWeight2 = 0
Form_Main.totalWeight2 = 0

Form_Main.pickerSESA = ""
Form_Main.pickerName = ""

Form_Main.Label100.ForeColor = RGB(0, 0, 0)
Form_Main.loadingWeight2.ForeColor = RGB(0, 0, 0)
Form_Main.Label105.ForeColor = RGB(0, 0, 0)
Form_Main.tareWeight2.ForeColor = RGB(0, 0, 0)
Form_Main.Label136.ForeColor = RGB(0, 0, 0)
Form_Main.totalWeight2.ForeColor = RGB(0, 0, 0)
Form_Main.Label140.ForeColor = RGB(0, 0, 0)

Form_Main.Label6.ForeColor = RGB(0, 0, 0)
Form_Main.loadingWeight.ForeColor = RGB(0, 0, 0)
Form_Main.Label90.ForeColor = RGB(0, 0, 0)
Form_Main.tareWeight.ForeColor = RGB(0, 0, 0)
Form_Main.Label135.ForeColor = RGB(0, 0, 0)
Form_Main.totalWeight.ForeColor = RGB(0, 0, 0)
Form_Main.Label139.ForeColor = RGB(0, 0, 0)

Form_Main.Frame142.Visible = False
Form_Main.Label143.Visible = False
Form_Main.Label152.Visible = False
Form_Main.Label156.Visible = False
Form_Main.Text151.Visible = False
Form_Main.Label157.Visible = False
Form_Main.Label153.Visible = False
Form_Main.GoCheck.Visible = False
Form_Main.SSCCNumber.Locked = False

End Sub

Public Sub initiateAndon()

Dim sql2 As String
Dim db As DAO.Database
Set db = CurrentDb

sql2 = "Delete * FROM [handlingUnitDetails] WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
db.Execute sql2

Form_Main.SSCCNumber = ""
Form_Main.SSCCnumber2 = ""
Form_Main.HandlingUnitsSubform.Requery

'TODO #10
'WaitSeconds (0.5)
Form_Main.orangeLight.Visible = True
Form_Main.greenLight.Visible = True
Form_Main.redLight.Visible = True

End Sub

Public Sub displayRedLight()

sndPlaySound32 "C:\Windows\Media\ringout.wav", &H1
'TODO #10
'i = 0
'Do While i < 4
'WaitSeconds (0.2)
'If Me.redLight.Visible = False Then
'Me.redLight.Visible = True
'Me.Repaint
'Else
'Me.redLight.Visible = False
'Me.Repaint
'End If
'WaitSeconds (0.2)
'i = i + 1
'Loop

Form_Main.redLight.Visible = True
Form_Main.greenLight.Visible = False
Form_Main.orangeLight.Visible = False
Form_Main.Repaint

End Sub

Public Sub displayOrangeLight()

sndPlaySound32 "C:\Windows\Media\notify.wav", &H1

'TODO #10
'i = 0
'Do While i < 4
'WaitSeconds (0.2)
'If Me.orangeLight.Visible = False Then
'Me.orangeLight.Visible = True
'Me.Repaint
'Else
'Me.orangeLight.Visible = False
'Me.Repaint
'End If
'WaitSeconds (0.2)
'i = i + 1
'Loop

Form_Main.redLight.Visible = False
Form_Main.greenLight.Visible = False
Form_Main.orangeLight.Visible = True
Form_Main.Repaint

End Sub


Public Sub displayGreenLight()

sndPlaySound32 "C:\Windows\Media\tada.wav", &H1
'TODO #10
'i = 0
'Do While i < 4
'WaitSeconds (0.2)
'If Me.greenLight.Visible = False Then
'Me.greenLight.Visible = True
'Me.Repaint
'Else
'Me.greenLight.Visible = False
'Me.Repaint
'End If
'WaitSeconds (0.2)
'i = i + 1
'Loop
Form_Main.redLight.Visible = False
Form_Main.greenLight.Visible = True
Form_Main.orangeLight.Visible = False
Form_Main.Repaint

End Sub

Public Function displaySAPweightData() As Boolean
If Form_Main.checkCalculatedWeight.Value = -1 Then
    displaySAPweightData = True
    Form_Main.Label6.ForeColor = RGB(216, 216, 216)
    Form_Main.loadingWeight.ForeColor = RGB(216, 216, 216)
    Form_Main.Label90.ForeColor = RGB(216, 216, 216)
    Form_Main.tareWeight.ForeColor = RGB(216, 216, 216)
    Form_Main.Label135.ForeColor = RGB(216, 216, 216)
    Form_Main.totalWeight.ForeColor = RGB(216, 216, 216)
    Form_Main.Label139.ForeColor = RGB(216, 216, 216)
    Exit Function
Else
    displaySAPweightData = False
    Form_Main.Label100.ForeColor = RGB(216, 216, 216)
    Form_Main.loadingWeight2.ForeColor = RGB(216, 216, 216)
    Form_Main.Label105.ForeColor = RGB(216, 216, 216)
    Form_Main.tareWeight2.ForeColor = RGB(216, 216, 216)
    Form_Main.Label136.ForeColor = RGB(216, 216, 216)
    Form_Main.totalWeight2.ForeColor = RGB(216, 216, 216)
    Form_Main.Label140.ForeColor = RGB(216, 216, 216)
    Exit Function
End If
End Function

Public Sub displayWCResults(result As String, diff As Double)

If result = "Excess" Then
    Form_Main.Label152.Caption = "Excess !"
    Form_Main.Label152.ForeColor = RGB(237, 28, 36)
    Form_Main.Label156.Caption = "+"
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Weight discrepancy, 100% check must be performed"
ElseIf result = "Short" Then
    Form_Main.Label152.Caption = "Short !"
    Form_Main.Label152.ForeColor = RGB(237, 28, 36)
    Form_Main.Label156.Caption = ""
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Weight discrepancy, 100% check must be performed"
ElseIf result = "lightExcess" Then
    Form_Main.Label152.Caption = "Light Excess !"
    Form_Main.Label152.ForeColor = RGB(237, 28, 36)
    Form_Main.Label156.Caption = "+"
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Light Weight discrepancy, 100% check must be performed"
ElseIf result = "lightShort" Then
    Form_Main.Label152.Caption = "Light Short !"
    Form_Main.Label152.ForeColor = RGB(237, 28, 36)
    Form_Main.Label156.Caption = ""
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Light Weight discrepancy, 100% check must be performed"
ElseIf result = "Pass" Then
    Form_Main.Label152.Caption = "Pass !"
    Form_Main.Label152.ForeColor = RGB(0, 149, 48)
    Form_Main.Label156.Caption = ""
    Form_Main.Text151.Value = diff
ElseIf result = "Picker" Then
    Form_Main.Label152.Caption = "Picker check !"
    Form_Main.Label152.ForeColor = RGB(255, 102, 0)
    Form_Main.Label156.Caption = ""
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Picker check, 100% check must be performed"
ElseIf result = "Sensitive" Then
    Form_Main.Label152.Caption = "Sensitive Material!"
    Form_Main.Label152.ForeColor = RGB(255, 102, 0)
    Form_Main.Label156.Caption = ""
    Form_Main.Text151.Value = diff
    Form_Main.Label153.Visible = True
    Form_Main.GoCheck.Visible = True
    Form_Main.SSCCNumber.Locked = True
    Form_Main.Label153.Caption = "Sensitive Material check, 100% check must be performed"
End If

Form_Main.Frame142.Visible = True
Form_Main.Label143.Visible = True
Form_Main.Label152.Visible = True
Form_Main.Label156.Visible = True
Form_Main.Label157.Visible = True
Form_Main.Text151.Visible = True
End Sub
