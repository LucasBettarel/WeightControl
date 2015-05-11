VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MDissuesListSubform"
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

Private Sub checkRemove_AfterUpdate()
Form_MDissuesListSubform.Requery
End Sub
