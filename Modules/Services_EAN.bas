Attribute VB_Name = "Services_EAN"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub recordMDissue()

DoCmd.OpenForm "reportMDissues1", acNormal

Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset
Dim sql3 As String
Dim workRS3 As Recordset
Dim sql4 As String
Dim workRS4 As Recordset

DoCmd.SetWarnings False

sql = "DELETE * FROM [MDissuesList]"
DoCmd.RunSQL sql

sql2 = "Select * " & _
        "FROM [HandlingUnitDetails] " & _
        "WHERE [Status] <> 'complete'" & _
        "AND [SSCC] = '" & Form_Main.SSCCnumber2 & "';"
       
Set workRS2 = CurrentDb.OpenRecordset(sql2)

If workRS2.RecordCount > 0 Then

sql3 = "INSERT INTO [MDissuesList]" & _
        "SELECT [HandlingUnitDetails].[material],[HandlingUnitDetails].[description],[HandlingUnitDetails].[EAN13],[HandlingUnitDetails].[Unit quantity],[HandlingUnitDetails].[Unit weight],[HandlingUnitDetails].[EAN14],[HandlingUnitDetails].[Package quantity],[HandlingUnitDetails].[Package weight] " & _
        "FROM [HandlingUnitDetails] " & _
        "WHERE [Status] <> 'complete' " & _
        "AND [SSCC] = '" & Form_Main.SSCCnumber2 & "';"
DoCmd.RunSQL sql3

Form_MDissuesListSubform.Requery

Else

MsgBox "No issue to report"

End If

sql4 = "Select * " & _
        "FROM [MDissuesList] " & _
        "WHERE [removeCheck] = 0;"

Form_MDissuesListSubform.RecordSource = sql4
Form_MDissuesListSubform.Requery

DoCmd.SetWarnings True

End Sub

Public Sub initiateEANcheck()
Dim workRS3 As Recordset
Dim sql3 As String
Dim sql As String

setMainViews ("focusEAN")

DoCmd.SetWarnings False
sql = "DELETE * FROM [excessIssues]"
DoCmd.RunSQL sql

Form_Main.excessIssuesSubform.Requery

sql3 = "UPDATE [handlingUnitDetails]" & _
"SET [status] = ''," & _
"[Manual Check] = 0 ," & _
"[quantityChecked] = 0 ; "

DoCmd.RunSQL sql3
Form_Main.HandlingUnitsSubform.Requery
DoCmd.SetWarnings True

DoCmd.OpenForm "EAN13check", acNormal
Form_EAN13check.Move 7000, 5000

End Sub

Public Sub recordPickingIssues()
'TODO #6
'once validated, add picking issue to table (check if already exist) and that's it
'other methods to create :
'- validation form
'- validation method
'- call supervisor
'
'excess material table should not exist anymore
Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset
Dim sql3 As String
Dim workRS3 As Recordset
checkHU = False
checkExcess = False

sql = "Select * " & _
        "FROM [pickingIssuesReport] " & _
        "WHERE [SSCC] ='" & Form_Main.SSCCnumber2 & "';"
       
Set workRS = CurrentDb.OpenRecordset(sql)

If workRS.RecordCount > 0 Then
     result = MsgBox("Issues for this HU have already been recorded, do you want to update it", vbYesNo)
    If result = vbNo Then
        'dont record
        Exit Sub
    Else
        'update records
        sql = "DELETE * " & _
                "FROM [pickingIssuesReport] " & _
                "WHERE [SSCC] ='" & Form_Main.SSCCnumber2 & "';"
        DoCmd.RunSQL sql
    End If
End If

Set Data = CurrentDb
Set checkLog = Data.OpenRecordset("pickingIssuesReport")
sql2 = "Select * " & _
        "FROM [HandlingUnitDetails] " & _
        "WHERE [Status] <> 'complete'" & _
         "AND [SSCC] = '" & Form_Main.SSCCnumber2 & "';"
Set workRS2 = CurrentDb.OpenRecordset(sql2)

If workRS2.RecordCount > 0 Then
    
    workRS2.MoveFirst
    checkHU = True
    Do While workRS2.EOF = False
        checkLog.AddNew
        checkLog("SSCC").Value = Form_Main.SSCCnumber2
        checkLog("Picker SESA").Value = Form_Main.pickerSESA
        checkLog("Picker Name").Value = Form_Main.pickerName
        checkLog("Checking Date").Value = Date
        checkLog("Checking Time").Value = Time
        checkLog("Issue").Value = workRS2![status]
        checkLog("Material").Value = workRS2![material]
        checkLog("Quantity").Value = workRS2![quantity]
        checkLog("Actual Quantity").Value = workRS2![quantityChecked]
        checkLog("Comments").Value = ""
        checkLog.Update
        workRS2.MoveNext
    Loop
End If

'**********check excess material table

sql3 = "Select * " & _
        "FROM [excessIssues] ;"
Set workRS3 = CurrentDb.OpenRecordset(sql3)

If workRS3.RecordCount > 0 Then
    workRS3.MoveFirst
    Do While workRS3.EOF = False
        checkExcess = True
        checkLog.AddNew
        checkLog("SSCC").Value = Form_Main.SSCCnumber2
        checkLog("Picker SESA").Value = Form_Main.pickerSESA
        checkLog("Picker Name").Value = Form_Main.pickerName
        checkLog("Checking Date").Value = Date
        checkLog("Checking Time").Value = Time
        checkLog("Issue").Value = "other material"
        checkLog("Material").Value = workRS3![excessMaterial]
        checkLog("Quantity").Value = 0
        checkLog("Actual Quantity").Value = workRS3![excessQuantity]
        checkLog("Comments").Value = ""
        checkLog.Update
        workRS3.MoveNext
    Loop
End If

If checkHU = True Or checkExcess = True Then
    recordCheckResult "EAN", "Fail", 0, 0, "", 0, ""
    MsgBox "Issues for this HU have been successfully recorded, Please call a supervisor to take actions"
    setMainViews ("completed_test")
End If

End Sub

Public Sub recordExcessMaterial(excessEANfound As String)

Dim sql As String
Dim RS As Recordset

Dim sql2 As String
Dim RS2 As Recordset

Dim sql3
Dim rs3 As Recordset

Set db = CurrentDb

sql2 = "SELECT * FROM [Material Master] " & _
        "WHERE [EAN13] = '" & excessEANfound & "';"
    
Set RS2 = db.OpenRecordset(sql2)

If (RS2.RecordCount > 0 And IsNull(RS2.RecordCount) = False) Then
 
 excessMat = RS2![material]
 excessQty = RS2![Quantity1]
 excessDescr = RS2![description]
 
 Else

sql3 = "SELECT * FROM [Material Master] " & _
        "WHERE [EAN14] = '" & excessEANfound & "';"
    
Set rs3 = db.OpenRecordset(sql3)

If (rs3.RecordCount > 0 And IsNull(rs3.RecordCount) = False) Then
 
 excessMat = rs3![material]
 excessQty = rs3![Quantity2]
 excessDescr = rs3![description]
Else

MsgBox "Material not found, please record issue manually"
Exit Sub
End If

End If
 

sql = "SELECT * FROM [excessIssues] WHERE [excessEAN] = '" & excessEANfound & "';"

Set RS = db.OpenRecordset(sql)
   
If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then

 RS.Edit
 RS![excessQuantity] = RS![excessQuantity] + excessQty
 RS.Update
 
Else

Set Data = CurrentDb
Set checkLog = Data.OpenRecordset("excessIssues")

         checkLog.AddNew
         checkLog("excessEAN").Value = excessEANfound
         checkLog("excessMaterial").Value = excessMat
         checkLog("excessDescription").Value = excessDescr
         checkLog("excessQuantity").Value = excessQty
         checkLog.Update


End If

Form_Main.excessIssuesSubform.Requery
Form_HandlingUnitsSubform.Requery
Form_EAN13check.Command9.SetFocus
Form_EAN13check.EAN13scan.SetFocus

End Sub

Public Function checkHUok()

checkHUok = True

Dim sal As String
Dim RS As Recordset
Dim db As DAO.Database
Set db = CurrentDb


sql = "SELECT * FROM [HandlingUnitDetails] where [SSCC] ='" & Form_Main.SSCCnumber2 & "';"
Set RS = db.OpenRecordset(sql)

If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then

RS.MoveFirst

Do While RS.EOF = False

If RS![quantityChecked] < RS![quantity] Then
    RS.Edit
    RS![status] = "short"
    RS.Update
    checkHUok = False
End If

If RS![quantityChecked] > RS![quantity] Then
    RS.Edit
    RS![status] = "excess"
    RS.Update
    checkHUok = False
End If

If RS![quantityChecked] = RS![quantity] Then
    RS.Edit
    RS![status] = "complete"
    RS.Update
    
End If

RS.MoveNext
Loop

End If

Form_HandlingUnitsSubform.Requery

End Function

Public Sub ValidateEANCheck()
On Error GoTo Routine_Error

test = checkHUok

If test = True Then

MsgBox "All ok"

recordCheckResult "EAN", "Pass", 0, 0, "", 0, ""

If Form_Main.EANautomaticprinting = -1 Then
Dim sql2 As String
Dim workRS2 As Recordset

sql2 = "Select * " & _
        "FROM [printers] " & _
        "WHERE [workstation] = '" & getWorkstation & "';"

Set workRS2 = CurrentDb.OpenRecordset(sql2)

If workRS2.RecordCount = 0 Then Exit Sub

selectedPrinter = workRS2![Printer]
Set Application.Printer = Application.Printers(selectedPrinter)

selectLabel = "EAN13CheckLabel"
DoCmd.SelectObject acReport, Trim(selectLabel), True

DoCmd.PrintOut , , , , 1
End If

DoCmd.Close acForm, "EAN13check"
Form_Main.SSCCNumber.SetFocus
initiateAndon

Else

MsgBox "There is issues in this handling unit, please check"
DoCmd.Close acForm, "EAN13check"

End If

Routine_Exit:
    Exit Sub

Routine_Error:
    MsgBox "The printer has not been found"
    Resume Routine_Exit
End Sub

Public Sub EANScanUpdate()
Dim sql As String
Dim RS As Recordset
Dim db As DAO.Database
Set db = CurrentDb


EAN13toCheck = Form_EAN13check.EAN13scan

If IsNull(EAN13toCheck) = True Then
Exit Sub
End If

'************CHECK EAN13 **************
sql = "SELECT * FROM [HandlingUnitDetails] WHERE [EAN13] = '" & EAN13toCheck & "';"

Set RS = db.OpenRecordset(sql)
   
   If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then
   RS.MoveLast
   If RS.RecordCount = 1 Then
   
   If RS![quantityChecked] >= RS![quantity] Then
   MsgBox " Excess Qty"
   End If
   
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + 1
    RS.Update
   
  Else
   
   RS.MoveFirst
   check = False

   Do While RS.EOF = False And check = False
   
    If RS![quantityChecked] < RS![quantity] Then
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + 1
    RS.Update
    check = True
    Else
     RS.MoveNext
    End If
    Loop
    
If check = False Then
RS.MoveLast
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + 1
    RS.Update
End If
  End If
  

Else

'************CHECK EAN14 **************

sql = "SELECT * FROM [HandlingUnitDetails] WHERE [EAN14] = '" & EAN13toCheck & "';"

Set RS = db.OpenRecordset(sql)
   
   If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then
   RS.MoveLast
   If RS.RecordCount = 1 Then
   
   If RS![quantityChecked] + RS![Package quantity] > RS![quantity] Then
   MsgBox " Excess Qty"
   End If
   
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + RS![Package quantity]
    RS.Update
   
  Else
   
   RS.MoveFirst
   check = False
toAdd = 0
toAdd = RS![Package quantity]

   Do While RS.EOF = False And check = False
   
    If RS![quantityChecked] + toAdd <= RS![quantity] Then
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + toAdd
    RS.Update
    check = True
    Else
    
    If RS![quantityChecked] < RS![quantity] Then
    RS.Edit
    toAdd = toAdd - RS![quantity] + RS![quantityChecked]
    RS![quantityChecked] = RS![quantity]
    RS.Update
    Else
    
    RS.MoveNext
    End If
    End If
    Loop
    
If check = False Then
RS.MoveLast
    RS.Edit
    RS![quantityChecked] = RS![quantityChecked] + toAdd
    RS.Update
End If
End If

Else
MsgBox "This material is not in the handling unit"
recordExcessMaterial (EAN13toCheck)

'add in the list
End If
End If

Form_HandlingUnitsSubform.Requery

Form_EAN13check.EAN13scan = ""
Form_EAN13check.EAN13scan.SetFocus

End Sub

Public Sub manualCheckUpdate()

Dim sql As String
Dim RS As Recordset
Dim db As DAO.Database
Set db = CurrentDb
checkAll = True

Form_HandlingUnitsSubform.manualCheck.Requery
Form_HandlingUnitsSubform.Requery

sql = "SELECT * FROM [HandlingUnitDetails] WHERE [SSCC] ='" & Form_Main.SSCCnumber2 & " ';"
Set RS = db.OpenRecordset(sql)

If (RS.RecordCount > 0 And IsNull(RS.RecordCount) = False) Then
    RS.MoveFirst
    Do While RS.EOF = False
        If RS![status] = "complete" Or RS![Manual Check].Value = -1 Then
            'line ok
        Else
            checkAll = False
        End If
        RS.MoveNext
    Loop
End If
Form_HandlingUnitsSubform.Requery
If checkAll = True Then
    DoCmd.OpenForm "manual check pop up", acNormal
End If

End Sub
