Attribute VB_Name = "Services_Login"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Function getUserName() As String
Dim currentUser As String
currentUser = CStr(Form_loginForm.userLogin.Value)
If Not IsNull(currentUser) Then
    getUserName = currentUser
Else
    getUserName = "Unknown user"
End If
Exit Function
End Function

Public Function getWorkstation() As String
Dim currentWS As String
currentWS = CStr(Form_loginForm.workstationName.Value)
If Not IsNull(currentWS) Then
    getWorkstation = currentWS
Else
    getWorkstation = "Workstation not defined"
End If
Exit Function
End Function


Public Sub checkWorstationParams()
Dim currentWS As String
Dim sql2 As String
Dim RS2 As Recordset

sql2 = "SELECT * FROM [printers] WHERE [Workstation] = '" & getWorkstation & "';"
Set RS2 = CurrentDb.OpenRecordset(sql2)

If RS2.RecordCount = 0 Then
MsgBox "Missing workstation parameters"
Exit Sub
End If

If RS2![Weight Control] = -1 Then
Form_Main.weightStationCheck = -1
Else
Form_Main.weightStationCheck = 0
End If

If RS2![EAN Quality Check] = -1 Then
Form_Main.EANstationCheck = -1
Else
Form_Main.EANstationCheck = 0
End If

cleaningWC

End Sub
Public Function isWC() As Boolean
isWC = False
If (Form_Main.weightStationCheck = -1) Or (Form_Main.weightStationCheck = 0 And Form_Main.EANstationCheck = 0) Then
    isWC = True
    Exit Function
End If
Exit Function
End Function

Public Function isEAN() As Boolean
isEAN = False
If (Form_Main.EANstationCheck = -1) Or (Form_Main.weightStationCheck = 0 And Form_Main.EANstationCheck = 0) Then
    isEAN = True
    Exit Function
End If
Exit Function
End Function

Public Function isADMIN() As Boolean
isADMIN = False
Dim sql As String
Dim RS As Recordset

sql = "SELECT * FROM [logonData] WHERE [SESA] = '" & getUserName & "';"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount <> 0 Then
    If RS![ADMIN] = -1 Then
        isADMIN = True
        Exit Function
    End If
End If
End Function

Public Sub setMainViews(action As String)
Dim WC As Boolean
Dim EAN As Boolean
Dim ADMIN As Boolean
WC = isWC
EAN = isEAN
ADMIN = isADMIN

If action = "login" Then
    If WC And Not EAN Then
        Form_Main.SSCCNumber.SetFocus
        Form_Main.Material_Control.Visible = False
    ElseIf Not WC And EAN Then
        Form_Main.SSCCnumber2.SetFocus
        Form_Main.Weight_Control.Visible = False
    Else
        Form_Main.SSCCNumber.SetFocus
    End If
    If ADMIN Then
        Form_Main.Settings.Visible = True
    Else
        Form_Main.Settings.Visible = False
    End If
End If

If action = "focusEAN" Then
    If Not ADMIN Then
       Form_Main.Weight_Control.Visible = False
        Form_Main.Quality_KPI.Visible = False
        Form_Main.Settings.Visible = False
        Form_Main.Calibration.Visible = False
    End If
Form_Main.Material_Control.Visible = True
Form_Main.Material_Control.SetFocus
End If

If action = "completed_test" Then
    If WC And Not EAN Then
        Form_Main.SSCCNumber.SetFocus
        Form_Main.Material_Control.Visible = False
        initiateAndon
        cleaningWC

    ElseIf Not WC And EAN Then
        Form_Main.SSCCnumber2.SetFocus
        Form_Main.Weight_Control.Visible = False
    Else
        Form_Main.SSCCNumber.SetFocus
        initiateAndon
        cleaningWC
    End If
    If ADMIN Then
        Form_Main.Settings.Visible = True
    Else
        Form_Main.Settings.Visible = False
    End If
End If



End Sub

Public Sub login_process()
Dim sql As String
Dim RS As Recordset
Dim AdminUser As Boolean
Dim OfflineUser As Boolean
AdminUser = isADMIN
OfflineUser = False

sql = "SELECT * FROM [logonData] WHERE [SESA] = '" & getUserName & "';"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount = 0 Then
    MsgBox "User unknown"
    Exit Sub
End If

If AdminUser And Form_loginForm.passwordLogin = RS![Password] Then OfflineUser = True

If Not OfflineUser Then
    rfc_call_logon
    If objConnection Is Nothing Then Exit Sub
Else
    MsgBox "Offline admin access"
End If

DoCmd.OpenForm "Main", acNormal

checkWorstationParams

setMainViews ("login")

End Sub
