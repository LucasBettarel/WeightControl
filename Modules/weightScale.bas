Attribute VB_Name = "weightScale"
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Antoine NICOLE
'Stephen HOUSSAYE
'Lucas BETTAREL


Option Compare Database

Public Sub initSerialPorts()

Dim lngStatus1 As Long
Dim strError1  As String

' Initialize Communications

Dim sql As String
Dim RS As Recordset

sql = "SELECT * FROM [printers] WHERE [Workstation] = '" & getWorkstation & "';"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount = 0 Then
MsgBox "No workstation selected"
Exit Sub
End If


    Dim scaleID As Integer
    Dim portCom As String
    
        scaleID = RS![scalePort]
        portCom = "COM" & scaleID
        
        lngStatus1 = CommOpen(scaleID, portCom, "baud=9600 parity=N data=8 stop=1")
    
        If lngStatus1 <> 0 Then
        'Handle error.
            lngStatus1 = CommGetError(strError1)
    
            'MsgBox "The weighting scale is not connected properly, please check"
        End If
    


End Sub

Public Sub ProcessDataFlow()

On Error GoTo PROC_ERR

Dim vData, vLastData As Long
Dim vPortID As Integer
Dim lngStatus3 As Long
Dim strError1  As String
Dim strData3 As String
Dim strReceive3 As String
Dim boxType As Long
Dim strLength3 As Long




Dim sql As String
Dim RS As Recordset

sql = "SELECT * FROM [printers] WHERE [Workstation] = '" & getWorkstation & "';"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount = 0 Then
MsgBox "No workstation selected"
Exit Sub
End If


    Dim scaleID As Integer
    Dim portCom As String
    
        scaleID = RS![scalePort]

check = 0

Do While check = 0
    DoEvents
    
    
    lngStatus3 = commRead(scaleID, strReceive3, 64)
    If lngStatus3 > 0 Then
       
        strData3 = strData3 + strReceive3
        strReceive3 = ""
        If InStr(strData3, Chr(10)) > 1 Or InStr(strData3, Chr(138)) > 1 Or Len(strData3) > 30 Then
            ' Process data.
            strData3 = Left(strData3, 30)
            strDataConv3 = ""
            'If InStr(strData3, "g") > 1 Then
            '    strData3 = Trim(Left(strData3, InStr(strData3, "g") - 2))
             '   Do While Not IsNumeric(strData3)
             '       strData3 = Trim(Right(strData3, Len(strData3) - 1))
                    
              '  Loop
              '  strLength3 = Len(strData3)
              '  For j = 1 To strLength3
               '     strDataConv3 = strDataConv3 + Chr(Asc(Left(strData3, 1)) Mod 128)
               '     strData3 = Mid(strData3, 2, 64)
            
              '  Next j
              '  Form_weightInput.weightMeasuredData = strDataConv3
              ' Form_Main.measuredWeight = strDataConv3
              ' check = 1
               
            'strData3 = ""
                
        'End If
      
            If InStr(strData3, "!") > 1 Then
            '    strData3 = Trim(Left(strData3, InStr(strData3, "g") - 2))
             '   Do While Not IsNumeric(strData3)
             '       strData3 = Trim(Right(strData3, Len(strData3) - 1))
                    
              '  Loop
              '  strLength3 = Len(strData3)
              '  For j = 1 To strLength3
               '     strDataConv3 = strDataConv3 + Chr(Asc(Left(strData3, 1)) Mod 128)
               '     strData3 = Mid(strData3, 2, 64)
            
              '  Next j
              
              strDataConv3 = Mid(strData3, InStr(strData3, "!") + 1, 6)
              strDataConv3 = Trim(strDataConv3)
               Form_weightInput.weightMeasuredData = strDataConv3
               Form_Main.measuredWeight = strDataConv3
               check = 1
               
            'strData3 = ""
                
        End If
      
      
 End If
 End If
 
Loop

PROC_EXIT:
  Exit Sub

PROC_ERR:
  'MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.WaitSeconds"
  Resume PROC_EXIT



End Sub

Public Sub closeSerialPorts()


    ' Close communications.
Dim sql As String
Dim RS As Recordset

sql = "SELECT * FROM [printers] WHERE [Workstation] = '" & getWorkstation & "';"
Set RS = CurrentDb.OpenRecordset(sql)

If RS.RecordCount = 0 Then
MsgBox "No workstation selected"
Exit Sub
End If


    Dim scaleID As Integer
  
    
        scaleID = RS![scalePort]
        
    Call CommClose(scaleID)



End Sub
