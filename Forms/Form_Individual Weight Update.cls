VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Individual Weight Update"
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

Private Sub Command10381_Click()

DoCmd.Close acForm, "Individual Weight Update"

End Sub


Private Sub Command16_Click()
On Error GoTo Routine_Error

DoCmd.OpenForm "weightInput", acNormal
'DoCmd.MoveSize,
Form_weightInput.Move 6000, 4000

 Call closeSerialPorts
 Call initSerialPorts
 Call ProcessDataFlow
 Call closeSerialPorts

DoCmd.Close acForm, "weightInput"


Me.MDweightInput = Form_Main.measuredWeight

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End

End Sub

Private Sub Command23_Click()

On Error GoTo Routine_Error
DoCmd.OpenForm "weightInput", acNormal
'DoCmd.MoveSize,
Form_weightInput.Move 6000, 4000

 Call closeSerialPorts
 Call initSerialPorts
 Call ProcessDataFlow
 Call closeSerialPorts

DoCmd.Close acForm, "weightInput"


Me.MDweightInput2 = Form_Main.measuredWeight

Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End
End Sub

Private Sub Command9_Click()

'save record

Dim sql As String
Dim RS As Recordset

If Me.MDmat = "" Or IsNull(Me.MDmat) = True Then Exit Sub

'check secondary table first
sql = "SELECT * FROM [Local Material Master] " & _
"WHERE [Material] ='" & Me.MDmat & "';"


Set RS = CurrentDb.OpenRecordset(sql)

If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then

    RS.Edit

RS![description] = Me.MDdescr
RS![material] = Me.MDmat
RS![EAN13] = Me.MDEAN13
RS![Quantity1] = Me.MDqty
RS![Weight1] = Me.MDweightInput
RS![EAN14] = Me.MDEAN14
RS![Quantity2] = Me.MDqty2
RS![Weight2] = Me.MDweightInput2
RS![Record Date] = Date
    RS.Update


Else

Set Data = CurrentDb
Set checkLog = Data.OpenRecordset("Local Material Master")

   checkLog.AddNew
   
checkLog("Description").Value = Me.MDdescr
checkLog("material").Value = Me.MDmat
checkLog("EAN13").Value = Me.MDEAN13
checkLog("Quantity1").Value = Me.MDqty
checkLog("Weight1").Value = Me.MDweightInput
checkLog("EAN14").Value = Me.MDEAN14
checkLog("Quantity2").Value = Me.MDqty2
checkLog("Weight2").Value = Me.MDweightInput2
checkLog("Record Date").Value = Date
checkLog.Update
End If





MDcleaning
Me.MDmat = ""
Me.MDEANScan = ""

Me.MDEANScan.SetFocus


End Sub




Private Sub MDEANScan_AfterUpdate()

MDcleaning
Me.MDmat = ""

If Len(Me.MDEANScan) <> 13 And Len(Me.MDEANScan) <> 14 Then
MsgBox "The bar code scanned is incorrect"
Exit Sub
End If

Dim sql As String
Dim RS As Recordset
Dim sql2 As String
Dim RS2 As Recordset
Dim db As DAO.Database
Set db = CurrentDb

'check secondary table first
sql = "SELECT * FROM [Local Material Master] " & _
"WHERE [EAN13] ='" & Me.MDEANScan & "'" & _
"OR [EAN14] ='" & Me.MDEANScan & "';"

Set RS = db.OpenRecordset(sql)

If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then

Me.MDdescr = RS![description]
Me.MDmat = RS![material]
Me.MDEAN13 = RS![EAN13]
Me.MDqty = RS![Quantity1]
Me.MDweightInput = RS![Weight1]
Me.MDEAN14 = RS![EAN14]
Me.MDqty2 = RS![Quantity2]
Me.MDweightInput2 = RS![Weight2]

Else

'check material master seconds

sql2 = "SELECT * FROM [Material Master] " & _
"WHERE [EAN13] ='" & Me.MDEANScan & "'" & _
"OR [EAN14] ='" & Me.MDEANScan & "';"

Set RS2 = db.OpenRecordset(sql2)

If (RS2.RecordCount > 0 And IsNull(RS2.RecordCount) = False) Then

Me.MDdescr = RS2![description]
Me.MDmat = RS2![material]
Me.MDEAN13 = RS2![EAN13]
Me.MDqty = RS2![Quantity1]
Me.MDweightInput = RS2![Weight1]
Me.MDEAN14 = RS2![EAN14]
Me.MDqty2 = RS2![Quantity2]
Me.MDweightInput2 = RS2![Weight2]


End If
End If




End Sub

Private Sub MDmat_AfterUpdate()

MDcleaning
Me.MDEANScan = ""

Dim sql As String
Dim RS As Recordset
Dim sql2 As String
Dim RS2 As Recordset
Dim db As DAO.Database
Set db = CurrentDb

'check secondary table first
sql = "SELECT * FROM [Local Material Master] " & _
"WHERE [Material] ='" & Me.MDmat & "';"


Set RS = db.OpenRecordset(sql)

If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then

Me.MDdescr = RS![description]
Me.MDmat = RS![material]
Me.MDEAN13 = RS![EAN13]
Me.MDqty = RS![Quantity1]
Me.MDweightInput = RS![Weight1]
Me.MDEAN14 = RS![EAN14]
Me.MDqty2 = RS![Quantity2]
Me.MDweightInput2 = RS![Weight2]

Else

'check material master seconds

sql2 = "SELECT * FROM [Material Master] " & _
"WHERE [Material] ='" & Me.MDmat & "';"

Set RS2 = db.OpenRecordset(sql2)

If (RS2.RecordCount > 0 And IsNull(RS2.RecordCount) = False) Then

Me.MDdescr = RS2![description]
Me.MDmat = RS2![material]
Me.MDEAN13 = RS2![EAN13]
Me.MDqty = RS2![Quantity1]
Me.MDweightInput = RS2![Weight1]
Me.MDEAN14 = RS2![EAN14]
Me.MDqty2 = RS2![Quantity2]
Me.MDweightInput2 = RS2![Weight2]

End If
End If




End Sub

Sub MDcleaning()


Me.MDdescr = ""

Me.MDEAN13 = ""
Me.MDqty = 0
Me.MDweightInput = 0
Me.MDEAN14 = ""
Me.MDqty2 = 0
Me.MDweightInput2 = 0


End Sub
