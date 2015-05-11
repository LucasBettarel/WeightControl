Attribute VB_Name = "BRIDGE SAP"
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Antoine NICOLE
'Stephen HOUSSAYE
'Lucas BETTAREL


Option Compare Database

Public objConnection, funcControl, conn As Object
Public RFC_READ_TABLE, tblOptions, tblData, tblFields, strExport1, strExport2 As Object
Public RfcCallTransaction, bdcTable As Object
Public result, SSCC, Response, Language, vTO As String

Public j, k As Long
Public test As Boolean


Public Sub rfc_call_logon()

'Sap opening
Set funcControl = CreateObject("SAP.Functions")
Set conn = funcControl.Connection

test = conn.IsConnected
If Not test Then
    Set objConnection = getSapConnection
    If objConnection Is Nothing Then Exit Sub

    If Not objConnection Is Nothing Then
        funcControl.Connection = objConnection
        'Applying RFC connection
        Set RFC_READ_TABLE = funcControl.Add("RFC_READ_TABLE")
        Set strExport1 = RFC_READ_TABLE.exports("QUERY_TABLE")
        Set strExport2 = RFC_READ_TABLE.exports("DELIMITER")
        Set tblOptions = RFC_READ_TABLE.Tables("OPTIONS")
        Set tblData = RFC_READ_TABLE.Tables("DATA")
        Set tblFields = RFC_READ_TABLE.Tables("FIELDS")
    
        Set RfcCallTransaction = funcControl.Add("RFC_CALL_TRANSACTION_USING")
        Set bdcTable = RfcCallTransaction.Tables("BDCTABLE")
       ' RfcCallTransaction.exports("UPDMODE") = "S"
    RfcCallTransaction.exports("MODE") = "A"
        test = True
       
    
    Else
        Response = MsgBox("No SAP connection", vbOKOnly)

    End If
    
End If
    
End Sub

Private Function getSapConnection() As Object

Dim objFileSystemObject As Object
Dim ctlLogon As Object
Dim ctlTableFactory As Object
Dim objWindowsScriptShell As Object
Dim objConnection As Object


Dim sql As String
Dim RS As Recordset

sql = "SELECT * FROM [parameters]"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount = 0 Then
MsgBox "BRIDGE parameters missing, please check"
Exit Function
End If


bridgeServer = RS![server]
bridgeSystem = RS![system]
bridgeClient = RS![client]
bridgeLanguage = RS![Language]
bridgeSystemNumber = RS![system number]

    Set oBapiCtrl = CreateObject("sap.bapi.1")
    Set oBapiLogon = CreateObject("sap.logoncontrol.1")

    oBapiCtrl.Connection = oBapiLogon.NewConnection
    oBapiCtrl.Connection.ApplicationServer = bridgeServer
    oBapiCtrl.Connection.system = bridgeSystem
    oBapiCtrl.Connection.client = bridgeClient
    oBapiCtrl.Connection.USER = getUserName
    oBapiCtrl.Connection.Password = Form_loginForm.passwordLogin
    oBapiCtrl.Connection.Language = bridgeLanguage
    oBapiCtrl.Connection.SystemNumber = bridgeSystemNumber
    
    Set conn = funcControl.Connection
    
   If oBapiCtrl.Connection.Logon(0, True) <> True Then
    MsgBox "not connected", vbInformation, "SAP Logon"
    Exit Function
    Else
    Set getSapConnection = oBapiCtrl.Connection
   End If


End Function

Public Sub RFC_CALL_TRANSACTION_USING()

Dim textLine As String
Dim vHandle, vDelivery, vHU As String
Dim vDatas, vLastDatas, vItem As Integer

If IsEmpty(funcControl) Then
    Call rfc_call_logon
   
ElseIf funcControl.Connection.IsConnected = 0 Then
    Call rfc_call_logon
    
Else
    
End If


'Load handling unit weight table

strExport1.Value = "VEKP"
strExport2.Value = "@"

tblFields.AppendRow
tblFields(1, "FIELDNAME") = "EXIDV"

tblFields.AppendRow
tblFields(2, "FIELDNAME") = "HANDLE"
    
tblFields.AppendRow
tblFields(3, "FIELDNAME") = "VENUM"

tblFields.AppendRow
tblFields(4, "FIELDNAME") = "BRGEW"

tblFields.AppendRow
tblFields(5, "FIELDNAME") = "NTGEW"

tblFields.AppendRow
tblFields(6, "FIELDNAME") = "TARAG"

tblFields.AppendRow
tblFields(7, "FIELDNAME") = "VHILM"

tblOptions.AppendRow

'HANDLE - Int.ID
'VENUM - Internal HU
'BRGEW - Total weight
'NTGEW - Loading weight
'TARAG - Tare weight
'VHILM   packaging material


tblOptions(1, "TEXT") = "EXIDV EQ '00" & Form_Main.SSCCNumber & "'"
    
If RFC_READ_TABLE.Call = True Then
    If tblData.RowCount > 0 Then

        textLine = tblData(1, "WA")
        vHU = getPieceOfText(textLine)
        vHandle = getPieceOfText(textLine)
        vInternalHU = getPieceOfText(textLine)
        vTotalWeight = getPieceOfText(textLine)
        vLoadWeight = getPieceOfText(textLine)
        vTare = getPieceOfText(textLine)
        vPackaging = getPieceOfText(textLine)
       
       If Len(vHU) = 20 Then vHU = Right(vHU, 18)
       
       
        tblData.freetable
        tblFields.freetable
    
    
        
        With Form_Main
        
        .totalWeight = vTotalWeight * 1000
        .loadingWeight = vLoadWeight * 1000
        .tareWeight = vTare * 1000
        .packagingType = vPackaging
       
        End With
        
  End If
 End If
 
 
 'Load handling unit material table

strExport1.Value = "VEPO"
strExport2.Value = "@"


tblFields.AppendRow
tblFields(1, "FIELDNAME") = "VENUM"
    
tblFields.AppendRow
tblFields(2, "FIELDNAME") = "VBELN"

tblFields.AppendRow
tblFields(3, "FIELDNAME") = "POSNR"

tblFields.AppendRow
tblFields(4, "FIELDNAME") = "VEMNG"

tblFields.AppendRow
tblFields(5, "FIELDNAME") = "MATNR"

tblOptions.AppendRow

'VENUM - Internal HU
'VBELN -Delivery
'POSNR -Item
'VEMNG - Packed quantity
'MATNR -Material


tblOptions(1, "TEXT") = "VENUM EQ '" & vInternalHU & "'"
    
Dim sql2 As String
Dim db As DAO.Database
Set db = CurrentDb
Dim sql As String

sql2 = "Delete * FROM [handlingUnitDetails] WHERE [SSCC] = '" & vHU & "';"
db.Execute sql2
    
If RFC_READ_TABLE.Call = True Then
    If tblData.RowCount > 0 Then

        vLastDatas = tblData.RowCount
        For vDatas = 1 To vLastDatas
        
        textLine = tblData(vDatas, "WA")
        vInternalHU = getPieceOfText(textLine)
        vDelivery = getPieceOfText(textLine)
        vDeliveryItem = getPieceOfText(textLine)
        vQuantity = getPieceOfText(textLine)
        vMaterial = getPieceOfText(textLine)
              
sql = "INSERT INTO [handlingUnitDetails]([SSCC],[internalHU],[delivery],[deliveryItem],[quantity],[material]) " & _
      "VALUES ('" & vHU & "','" & vInternalHU & "','" & vDelivery & "','" & vDeliveryItem & "','" & vQuantity & "','" & vMaterial & "');"

db.Execute sql

Next vDatas
End If

        tblData.freetable
        tblFields.freetable
    
 End If
 
'lier VEPO <> LTAP
'lier delivery + Item <> VBELN + POSNR
'Champs picker = ENAME
Dim deliveryID As String
deliveryID = vDelivery & vDeliveryItem
 
strExport1.Value = "LTAP"
strExport2.Value = "@"

tblFields.AppendRow
tblFields(1, "FIELDNAME") = "VBELN"

tblFields.AppendRow
tblFields(2, "FIELDNAME") = "POSNR"

tblFields.AppendRow
tblFields(3, "FIELDNAME") = "ENAME"


tblOptions.AppendRow


'VBELN -Delivery
'POSNR -Item
'ENAME - PICKER SESA



If IsNull(vDelivery) = True Or IsEmpty(vDelivery) Then
MsgBox "This HU is empty, please check"
Exit Sub
End If



tblOptions(1, "TEXT") = "VBELN EQ '" & vDelivery & "'"
    
vSESA = ""

If RFC_READ_TABLE.Call = True Then
    If tblData.RowCount > 0 Then

        vLastDatas = tblData.RowCount
        For vDatas = 1 To vLastDatas
        
        textLine = tblData(vDatas, "WA")
        
        vDelivery = getPieceOfText(textLine)
        vDeliveryItem = getPieceOfText(textLine)
        vSESA = getPieceOfText(textLine)
If deliveryID = vDelivery & vDeliveryItem Then
Form_Main.pickerSESA = vSESA
Exit For
End If

Next vDatas
End If

        tblData.freetable
        tblFields.freetable
 End If
 
Dim sql4 As String
Dim workRS4 As Recordset

sql4 = "Select * " & _
        "FROM [pickerList] " & _
        "WHERE [SESA] = '" & vSESA & "';"

Set workRS4 = CurrentDb.OpenRecordset(sql4)

If workRS4.RecordCount > 0 Then

Form_Main.pickerName = workRS4![Picker Name]
Form_Main.picker100check = workRS4![100%]

End If

 
DoCmd.SetWarnings False

Dim sql9 As String
Dim workRS9 As Recordset

sql9 = "Select * " & _
        "FROM [handlingUnitDetails] " & _
        "INNER JOIN [Material Master] ON [handlingUnitDetails].Material = [Material Master].Material ;"

Set workRS9 = CurrentDb.OpenRecordset(sql9)

Dim sql3 As String

sql3 = "UPDATE [handlingUnitDetails]" & _
"INNER JOIN [Material Master] ON [handlingUnitDetails].Material = [Material Master].Material " & _
"SET [handlingUnitDetails].[Unit quantity] = [Material Master].[Quantity1]," & _
"[handlingUnitDetails].[Unit Weight] = [Material Master].[Weight1], " & _
"[handlingUnitDetails].[description] = [Material Master].[Description], " & _
"[handlingUnitDetails].[EAN14] = [Material Master].[EAN14], " & _
"[handlingUnitDetails].[Package quantity] = [Material Master].[Quantity2], " & _
"[handlingUnitDetails].[EAN13] = [Material Master].[EAN13], " & _
"[handlingUnitDetails].[Package Weight] = [Material Master].[Weight2]; "

DoCmd.RunSQL sql3

sql3 = "UPDATE [handlingUnitDetails]" & _
"INNER JOIN [sensitiveMaterial] ON [handlingUnitDetails].Material = [sensitiveMaterial].Material " & _
"SET [handlingUnitDetails].[Sensitive Material] = [sensitiveMaterial].[sensitive];"

DoCmd.RunSQL sql3


If workRS9.RecordCount > 0 Then
    If Form_Main.checkCalculatedWeight.Value = 0 Then
        Form_Main.checkSAPWeight.Value = -1
        Form_Main.checkCalculatedWeight.Value = 0
    Else
        Form_Main.checkSAPWeight.Value = 0
        Form_Main.checkCalculatedWeight.Value = -1
    End If
End If

Dim sql5 As String
Dim workRS5 As Recordset

sql5 = "Select * " & _
        "FROM [handlingUnitDetails] " & _
        "WHERE [Sensitive Material] = -1 ;"

Set workRS5 = CurrentDb.OpenRecordset(sql5)

If workRS5.RecordCount > 0 Then
DoCmd.OpenForm "Main", acNormal
Form_Main.sensitiveCheck = -1

End If

DoCmd.SetWarnings True
 

End Sub
    

Private Function getPieceOfText(textLine As String) As String

Dim position As Long

If Right(textLine, 1) <> "@" Then textLine = textLine & "@"

position = InStr(1, textLine, "@")
getPieceOfText = CStr(Left(textLine, position - 1))
textLine = Right(textLine, Len(textLine) - position)

End Function


