VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Private Sub Command12_Click()
DoCmd.Close
End Sub

Private Sub Form_Load()
 If Form_Main.TabCtl0.Value = 4 Then
    Me.img_calib.Visible = True
    Me.Img_help1.Visible = False
    Me.Img_help2.Visible = False
    Me.InsideHeight = 596 * 13
Else
    Me.img_calib.Visible = False
    Me.Img_help1.Visible = True
    Me.Img_help2.Visible = True
    Me.InsideHeight = 596 * 13
 End If
End Sub
