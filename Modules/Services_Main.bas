Attribute VB_Name = "Services_Main"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Declare Function sndPlaySound32 _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Public Sub WaitSeconds(intSeconds As Integer)
  ' Comments: Waits for a specified number of seconds
  ' Params  : intSeconds      Number of seconds to wait
  ' Source  : Total Visual SourceBook

  On Error GoTo PROC_ERR

  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
   ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.description, , "modDateTime.WaitSeconds"
  Resume PROC_EXIT
End Sub

Public Sub recordCheckResult(checkType As String, checkResult As String, theoryWeight As Double, actualWeight As Double, Comment As String, boxW As Double, boxT As String)
Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset
Dim sql3 As String
Dim workRS3 As Recordset

'TODO #10 : setmainviews redondant : check overall process
'If checkType = "EAN" Or checkType = "Manual" Then
'    If Form_Main.weightStationCheck = -1 Then
'        Form_Main.Weight_Control.Visible = True
'        Me.Quality_KPI.Visible = True
'        Me.Calibration.Visible = True
'    End If
'End If
'
'If isADMIN Then
'    Me.Settings.Visible = True
'End If

sql = "Select * " & _
        "FROM [QCcheckLog] " & _
        "WHERE [handlingUnit] ='" & Form_Main.SSCCNumber & "'" & _
        "AND [checkType] = '" & checkType & "';"
Set workRS = CurrentDb.OpenRecordset(sql)

If workRS.RecordCount > 0 Then
    result = MsgBox("This handling unit has already been tested, do you want to change result?", vbYesNo)
    If result = vbYes Then
        DoCmd.SetWarnings False
        sql = "Delete * " & _
                "FROM [QCcheckLog] " & _
                "WHERE [handlingUnit] ='" & Form_Main.SSCCNumber & "'" & _
                "AND [checkType] = '" & checkType & "';"
        DoCmd.RunSQL sql
        DoCmd.SetWarnings True
    Else
        Exit Sub
    End If
End If

sql2 = "SELECT * FROM [handlingUnitDetails] WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
Set workRS2 = CurrentDb.OpenRecordset(sql2)

If workRS2.RecordCount > 0 Then
    workRS2.MoveLast
    countLines = workRS2.RecordCount
End If

Set Data = CurrentDb
Set checkLog = Data.OpenRecordset("QCcheckLog")

    checkLog.AddNew
    checkLog("checkDate").Value = Date
    checkLog("checkTime").Value = Time
    checkLog("handlingUnit").Value = Form_Main.SSCCNumber
    checkLog("checkerName").Value = getUserName
    checkLog("Workstation").Value = getWorkstation
    checkLog("checkType").Value = checkType
    checkLog("Result").Value = checkResult
    If IsNull(theoryWeight) = False Then checkLog("TheoriticalWeight").Value = theoryWeight
    If IsNull(actualWeight) = False Then checkLog("ActualWeight").Value = actualWeight
    If IsNull(boxW) = False Then checkLog("boxWeight").Value = boxW
    checkLog("HUlines").Value = countLines
    checkLog("Comment").Value = Comment
    checkLog("boxType").Value = boxT
    checkLog.Update
            
 'delete item from waiting list after manuel/EAN check
 If checkType = "Manual" Or checkType = "EAN" Then
    DoCmd.SetWarnings False
     sql = "SELECT * FROM [waitingList] WHERE [SSCC] = '" & Form_Main.SSCCnumber2 & "';"
     Set workRS2 = CurrentDb.OpenRecordset(sql)
     If workRS2.RecordCount > 0 Then
       sql = "Delete * FROM [waitingList] WHERE [SSCC] = '" & Form_Main.SSCCnumber2 & "';"
       DoCmd.RunSQL sql
     End If
     DoCmd.SetWarnings True
End If
            
End Sub

Public Sub printCheckLabel(labelType As String)
If Form_Main.weightAutomaticPrinting = -1 Then
    Dim sql As String
    Dim workRS As Recordset
    sql = "Select * " & _
            "FROM [printers] " & _
            "WHERE [workstation] = '" & getWorkstation & "';"
    Set workRS = CurrentDb.OpenRecordset(sql)
    
    If workRS.RecordCount > 0 Then
        selectedPrinter = workRS![Printer]
        Set Application.Printer = Application.Printers(selectedPrinter)
        If labelType = "Weight" Then selectLabel = "weightCheckLabel"
        DoCmd.SelectObject acReport, Trim(selectLabel), True
        DoCmd.PrintOut , , , , 1
    End If
    
Routine_Exit:
        Exit Sub
    
Routine_Error:
        MsgBox "The printer has not been found"
        Resume Routine_Exit
End If
End Sub

Public Sub useLocalData()

Dim sql2 As String
Dim RS2 As Recordset
Dim sql As String
Dim RS As Recordset
Dim sql3 As String
Dim sql4 As String
Dim RS4 As Recordset

DoCmd.SetWarnings False

sql3 = "UPDATE [handlingUnitDetails]" & _
"INNER JOIN [Local Material Master] ON [handlingUnitDetails].Material = [Local Material Master].Material " & _
"SET [handlingUnitDetails].[Unit quantity] = iif(isnull([Local Material Master].[Quantity1]),[handlingUnitDetails].[Unit quantity],[Local Material Master].[Quantity1])," & _
"[handlingUnitDetails].[Unit Weight] = iif(isnull([Local Material Master].[Weight1]),[handlingUnitDetails].[Unit Weight],[Local Material Master].[Weight1]), " & _
"[handlingUnitDetails].[EAN14] = iif(isnull([Local Material Master].[EAN14]),[handlingUnitDetails].[EAN14], [Local Material Master].[EAN14]), " & _
"[handlingUnitDetails].[Package quantity] = iif(isnull([Local Material Master].[Quantity2]),[handlingUnitDetails].[Package quantity],[Local Material Master].[Quantity2]), " & _
"[handlingUnitDetails].[description] = iif(isnull([Local Material Master].[description]),[handlingUnitDetails].[description],[Local Material Master].[description]), " & _
"[handlingUnitDetails].[EAN13] = iif(isnull([Local Material Master].[EAN13]),[handlingUnitDetails].[EAN13],[Local Material Master].[EAN13]), " & _
"[handlingUnitDetails].[Local Masterdata] = -1 , " & _
"[handlingUnitDetails].[Package Weight] = iif(isnull([Local Material Master].[Weight2]),[handlingUnitDetails].[Package Weight],[Local Material Master].[Weight2]); "

DoCmd.RunSQL sql3
DoCmd.SetWarnings True
        
sql2 = "SELECT *" & _
       "FROM [handlingUnitDetails]" & _
       "WHERE [handlingUnitDetails].[SSCC] = '" & Form_Main.SSCCNumber & "'" & _
       "AND [handlingUnitDetails].[Local Masterdata] = -1 ;"
Set RS2 = CurrentDb.OpenRecordset(sql2)

'choose if recalculate all the time or only if new MD (now systematical)
If RS2.RecordCount > 0 Then
    Form_Main.checkCalculatedWeight.Value = -1
    Form_Main.checkSAPWeight = 0
End If

'recalculate weight
newWeight = 0
sql = "SELECT * FROM [handlingUnitDetails] WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
Set RS = CurrentDb.OpenRecordset(sql)
If RS.RecordCount = 0 Then Exit Sub
RS.MoveFirst

Do While RS.EOF = False
    If IsNull(RS![Unit quantity]) = True Or RS![Unit quantity] = 0 Or IsNull(RS![Unit weight]) = True Or RS![Unit weight] = 0 Then
        MsgBox "Incomplete masterdata, weight can't be recalculated"
        Exit Sub
    End If
    qty = RS![quantity]

    Do While qty > 0
        If IsNull(RS![Package quantity]) = False And RS![Package quantity] > 0 And RS![Package weight] > 0 And IsNull(RS![Package weight]) = False And (qty - RS![Package quantity]) >= 0 Then
            qty = qty - RS![Package quantity]
            newWeight = newWeight + RS![Package weight]
        Else
            qty = qty - RS![Unit quantity]
            newWeight = newWeight + RS![Unit weight]
        End If
    Loop
    
    RS.MoveNext
Loop

Form_Main.tareWeight2 = 0
Form_Main.loadingWeight2 = 0

sql4 = "SELECT * FROM [boxWeight] where [boxName] = '" & Form_Main.packagingType & "';"
Set RS4 = CurrentDb.OpenRecordset(sql4)

If RS4.RecordCount > 0 Then
    If RS4![emptyWeight] > 0 And IsNull(RS4![emptyWeight]) = False Then
        Form_Main.tareWeight2 = RS4![emptyWeight]
    Else
        If IsNull(Form_Main.tareWeight) = False Then Form_Main.tareWeight2 = Form_Main.tareWeight
    End If
Else
    If IsNull(Form_Main.tareWeight) = False Then Form_Main.tareWeight2 = Form_Main.tareWeight
End If

If IsNull(newWeight) = False Then Form_Main.loadingWeight2 = newWeight
Form_Main.totalWeight2.Value = Form_Main.tareWeight2.Value + Form_Main.loadingWeight2.Value

End Sub

Public Sub updateWaitingList()
Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset
Dim countLines As Double
Dim countHULines As Double

countLines = 0
countHULines = 0
            
'Check si deja enregistre dans la waiting list
sql = "Select * " & _
        "FROM [waitingList] " & _
        "WHERE [SSCC] ='" & Form_Main.SSCCnumber2 & "';"
Set workRS = CurrentDb.OpenRecordset(sql)

If workRS.RecordCount > 0 Then
    'compte nombre de lines et HUlines pour le sscc
    sql2 = "SELECT * FROM [handlingUnitDetails] WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
    Set workRS2 = CurrentDb.OpenRecordset(sql2)
    If workRS2.RecordCount > 0 Then
        workRS2.MoveLast
        countLines = workRS2.RecordCount
        workRS2.MoveFirst
        Do While workRS2.EOF = False
            countHULines = countHULines + workRS2![HUlines]
            workRS2.MoveNext
        Loop
        Form_Main.Qnb.Caption = CStr(countHULines)
    End If
    
    'color background indicator
    If countHULines = 0 Then
        Form_Main.Box242.BackColor = RGB(0, 149, 48)
    ElseIf countHULines > 0 And countHULines < 10 Then
        Form_Main.Box242.BackColor = RGB(255, 153, 51)
    Else
        Form_Main.Box242.BackColor = RGB(255, 0, 0)
    End If
End If

Dim Yesterday As Date
Yesterday = (Date) - 1
'old boxes alert
sql = "Select * " & _
        "FROM [waitingList] " & _
        "WHERE [checkDate] <> #" & Date & "# " & _
        "AND [checkDate] <> #" & Yesterday & "#;"
Set workRS = CurrentDb.OpenRecordset(sql)
 
End Sub

Public Sub updateHUWeight()
Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset

sql = "Select * " & _
        "FROM [QCcheckLog] " & _
        "WHERE [handlingUnit] ='" & Form_Main.SSCCnumber2 & "'" & _
        "AND [checkType] = 'Weight'" & _
        "OR [checkType] = 'Light Weight Check';"
        
Set workRS = CurrentDb.OpenRecordset(sql)
If workRS.RecordCount > 0 Then
    If workRS![actualWeight] <> 0 Or Not IsNull(workRS![actualWeight]) Then
        Form_Main.boxWeight.Value = workRS![actualWeight]
    Else
        Form_Main.boxWeight.Value = "not recorded"
    End If
Else
    MsgBox ("This handling unit has not been tested on Weight Check Station")
    Form_Main.boxWeight.Value = "not recorded"
End If
End Sub

Public Sub addFeedback()

Dim StrSQL As String
Dim memo As String

If IsNull(Form_Feedback.fb_memo) = True Then
    MsgBox "Please insert a message !"
    Exit Sub
Else
    memo = removeSpecial(Form_Feedback.fb_memo)
    DoCmd.SetWarnings False
    StrSQL = " INSERT INTO [Feedback]([Date_fb],[Workstation],[User],[Comment]) VALUES " _
               & "('" & Form_Main.Date_calib & "', '" & getWorkstation & "', '" & getUserName & "', '" & memo & "');"
    DoCmd.RunSQL StrSQL
    DoCmd.SetWarnings True
End If

MsgBox "Thank you, your feedback will be reviewed soon!"
DoCmd.Close

End Sub

Function removeSpecial(sInput As String) As String
    Dim sSpecialChars As String
    Dim i As Long
    sSpecialChars = "\/:*?""<>|[];'"
    For i = 1 To Len(sSpecialChars)
        sInput = Replace$(sInput, Mid$(sSpecialChars, i, 1), " ")
    Next
    removeSpecial = sInput
End Function


