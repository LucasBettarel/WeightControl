VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_loginForm"
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

Private Sub Command1038_Click()
DoCmd.Quit
End Sub

Private Sub Command9_Click()

If IsNull(Me.workstationName) = True Then
    MsgBox "Please select a workstation"
    Exit Sub
End If

If IsNull(Me.passwordLogin) = True Then
    MsgBox "Please enter a password"
    Exit Sub
End If

If IsNull(Me.userLogin) = True Then
    MsgBox "Please enter a user"
    Exit Sub
End If

login_process

DoCmd.Minimize
Me.Visible = False

End Sub

Private Sub Form_Load()
DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.ShowToolbar "Menu Bar", acToolbarNo
CurrentDb.Properties("StartUpShowDBWindow") = False
End Sub
