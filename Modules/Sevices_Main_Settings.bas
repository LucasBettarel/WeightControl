Attribute VB_Name = "Sevices_Main_Settings"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub uploadMaterialMaster()
Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True
  
strFile = findAttachment("Select Material file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If

Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False

Do While c < 252
If .Cells(1, c) <> "" Then
.Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
End If
c = c + 1
Loop

    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "Material Master", "C:\temp Material.xls", True

Dim sql As String
sql = "DELETE * FROM [Material Master]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "Material Master", strFile, True
Kill "C:\temp Material.xls"
MsgBox "Material table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "Material Master", "C:\temp Material.xls", True
Kill "C:\temp Material.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub uploadQCcheckLog()

Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True
  
strFile = findAttachment("Select QC log file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
    MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If

Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False

Do While c < 252
    If .Cells(1, c) <> "" Then
        .Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
    End If
    c = c + 1
Loop
    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "QCcheckLog", "C:\temp QCcheckLog.xls", True

Dim sql As String
sql = "DELETE * FROM [QCcheckLog]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "QCcheckLog", strFile, True
Kill "C:\temp QCcheckLog.xls"
MsgBox "QCcheckLog table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "QCcheckLog", "C:\temp QCcheckLog.xls", True
Kill "C:\temp QCcheckLog.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub uploadPickers()
Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True

strFile = findAttachment("Select Picker file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If

Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False

Do While c < 252
If .Cells(1, c) <> "" Then
.Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
End If
c = c + 1
Loop
    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "pickerList", "C:\temp pickerList.xls", True

Dim sql As String
sql = "DELETE * FROM [pickerList]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "pickerList", strFile, True
Kill "C:\temp pickerList.xls"
MsgBox "pickerList table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "pickerList", "C:\temp pickerList.xls", True
Kill "C:\temp pickerList.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub uploadSensitive()
Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True
  
strFile = findAttachment("Select sensitiveMaterial file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If

Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False
   
Do While c < 252
If .Cells(1, c) <> "" Then
.Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
End If
c = c + 1
Loop
    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "sensitiveMaterial", "C:\temp sensitiveMaterial.xls", True

Dim sql As String
sql = "DELETE * FROM [sensitiveMaterial]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "sensitiveMaterial", strFile, True
Kill "C:\temp sensitiveMaterial.xls"
MsgBox "pickerList table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "sensitiveMaterial", "C:\temp sensitiveMaterial.xls", True
Kill "C:\temp sensitiveMaterial.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub uploadPickingIssues()
Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True

strFile = findAttachment("Select pickingIssuesReport file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If

Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False
   
Do While c < 252
If .Cells(1, c) <> "" Then
.Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
End If
c = c + 1
Loop
    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "pickingIssuesReport", "C:\temp pickingIssuesReport.xls", True

Dim sql As String
sql = "DELETE * FROM [pickingIssuesReport]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "pickingIssuesReport", strFile, True
Kill "C:\temp pickingIssuesReport.xls"
MsgBox "pickingIssuesReport table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "pickingIssuesReport", "C:\temp pickingIssuesReport.xls", True
Kill "C:\temp pickingIssuesReport.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub uploadLocalMaterialMaster()
Dim strFile As String
Dim strFilter As String
Dim oXL As Object
Dim objWorkbook As Object

Set oXL = CreateObject("Excel.Application")

'Only XL 97 supports UserControl Property
On Error Resume Next
oXL.UserControl = True

strFile = findAttachment("Select Local Material Master file")

If strFile = "" Or IsNull(strFile) = True Then Exit Sub
If (Right(strFile, 4) <> ".xls") And (Right(strFile, 5) <> ".xlsx") Then
MsgBox ("The file " + strFile + " is not an excel file")
Exit Sub
End If
Set objWorkbook = oXL.Workbooks.Open(strFile)

With oXL
    .Application.DisplayAlerts = False
   
Do While c < 252
If .Cells(1, c) <> "" Then
.Cells(1, c) = Trim(Replace(.Cells(1, c), ".", " "))
End If
c = c + 1
Loop
    objWorkbook.saveAs strFile
    .Quit
End With

DoCmd.TransferSpreadsheet acExport, , "Local Material Master", "C:\temp Local Material Master.xls", True

Dim sql As String
sql = "DELETE * FROM [Local Material Master]"
DoCmd.RunSQL sql

DoCmd.TransferSpreadsheet acImport, , "Local Material Master", strFile, True
Kill "C:\temp Local Material Master.xls"
MsgBox "Local Material Master table has been successfully uploaded"
Exit Sub

cmdImport_Error:
DoCmd.TransferSpreadsheet acImport, , "Local Material Master", "C:\temp Local Material Master.xls", True
Kill "C:\temp Local Material Master.xls"
MsgBox "Something went wrong! Error was: " & Err.Number & " " & Err.description

End Sub

Public Sub LockApp()
If Form_Main.Loock = -1 Then
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    DoCmd.ShowToolbar "Menu Bar", acToolbarNo
    CurrentDb.Properties("StartUpShowDBWindow") = False
    Form_Main.unloockedImage.Visible = False
    Form_Main.loockedImage.Visible = True
 Else
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    DoCmd.ShowToolbar "Menu Bar", acToolbarYes
    CurrentDb.Properties("StartUpShowDBWindow") = True
    Form_Main.unloockedImage.Visible = True
    Form_Main.loockedImage.Visible = False
End If
End Sub
