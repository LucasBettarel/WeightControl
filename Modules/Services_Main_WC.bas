Attribute VB_Name = "Services_Main_WC"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub ValidateWCTest()
Dim sql2 As String
Dim workRS2 As Recordset
Dim countHULines As Double
Dim countLines As Double
countHULines = 0
countLines = 0

'count current hulines in waiting list
sql2 = "SELECT * FROM [waitingList];"
Set workRS2 = CurrentDb.OpenRecordset(sql2)
If workRS2.RecordCount > 0 Then
    workRS2.MoveLast
    countLines = workRS2.RecordCount
    workRS2.MoveFirst
    Do While workRS2.EOF = False
        countHULines = countHULines + workRS2![HUlines]
        workRS2.MoveNext
    Loop
End If

If isEAN Then
    initiateEANcheck
ElseIf countHULines > 10 Then
    'Overload -> weight control station perform ean check
    MsgBox "The EAN check station is currently overloaded. Please perform the EAN check here."
    initiateEANcheck
Else
    'weight control station normal fucntionment
    MsgBox "Handling Unit inserted into the EAN Station waiting list"
    setMainViews ("completed_test")
End If

End Sub

Public Function weightComparison() As Boolean

Dim resultComparison As Boolean
Dim measured As Double
Dim theoriticalWeight As Double
Dim boxWeight As Double
Dim boxType As String
Dim sql As String
Dim workRS As Recordset
Dim sql1 As String
Dim workRS1 As Recordset
Dim diffWeight As Double
Dim lightTolerance As Double
Dim lightest As Double

resultComparison = False
measured = Form_Main.measuredWeight
theoriticalWeight = 0
boxWeight = 0
boxType = ""
boxType = Form_Main.packagingType
lightTolerance = getLightTolerance
lightest = 0

'***************get theoritical weight
If displaySAPweightData Then
    theoriticalWeight = Form_Main.totalWeight2
    boxWeight = Form_Main.tareWeight2
Else
    theoriticalWeight = Form_Main.totalWeight
    boxWeight = Form_Main.tareWeight
End If
diffWeight = measured - theoriticalWeight

'*************get lightest item
sql = "Select MIN([Unit weight]) as minWeight " & _
        "FROM [HandlingUnitDetails] " & _
        "WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
Set workRS = CurrentDb.OpenRecordset(sql)
If workRS.RecordCount = 0 Then
    MsgBox "Min weight not found"
Else
    lightest = workRS![minWeight]
End If

'*************get weight tolerance
sql1 = "Select * " & _
        "FROM [weightControlSpecifications] " & _
        "WHERE [weightMin] <=" & theoriticalWeight & _
        "AND [weightMax] > " & theoriticalWeight & ";"
Set workRS1 = CurrentDb.OpenRecordset(sql1)
If workRS1.RecordCount = 0 Then
    MsgBox "There is no tolerance specifications, please check"
    Exit Function
End If

MAX = 0
min = 0
MAX = workRS1![positiveTolerance]
min = workRS1![negativeTolerance]

'***********LIGHT COMPARISON
If measured - theoriticalWeight >= k * lightest And Form_Main.lightestItem.Value = -1 Then
    displayWCResults "lightExcess", diffWeight
    recordCheckResult "Light Weight Check", "Excess", theoriticalWeight, measured, "", boxWeight, boxType
    addToQ "Excess", "Light Weight Check"
    weightComparison = resultComparison
    Exit Function
ElseIf theoriticalWeight - measured >= k * lightest And Form_Main.lightestItem.Value = -1 Then
    displayWCResults "lightShort", diffWeight
    recordCheckResult "Light Weight Check", "Short", theoriticalWeight, measured, "", boxWeight, boxType
    addToQ "Short", "Light Weight Check"
    weightComparison = resultComparison
    Exit Function

'******** WEIGHT COMPARISON
ElseIf measured > MAX + theoriticalWeight And Form_Main.weightCheck.Value = -1 Then
    displayWCResults "Excess", diffWeight
    recordCheckResult "Weight", "Excess", theoriticalWeight, measured, "", boxWeight, boxType
    addToQ "Excess", "Weight"
    weightComparison = resultComparison
    Exit Function
ElseIf measured < theoriticalWeight - min And Form_Main.weightCheck.Value = -1 Then
    displayWCResults "Short", diffWeight
    recordCheckResult "Weight", "Short", theoriticalWeight, measured, "", boxWeight, boxType
    addToQ "Short", "Weight"
    weightComparison = resultComparison
    Exit Function
    
'*******Pass
ElseIf measured >= (theoriticalWeight - min) And measured <= (theoriticalWeight + MAX) And Form_Main.weightCheck.Value = -1 Then
    displayWCResults "Pass", diffWeight
    resultComparison = True
    recordCheckResult "Weight", "Pass", theoriticalWeight, measured, "", boxWeight, boxType
    weightComparison = resultComparison
    Exit Function
Else
'******NO TEST DONE
    weightComparison = resultComparison
    Exit Function
End If
End Function

Public Function getLightTolerance() As Double
Dim k As Double
k = 0.5
Dim sql As String
Dim workRS As Recordset

sql = "Select * " & _
        "FROM [printers] " & _
        "WHERE [workstation] = '" & getWorkstation & "';"
Set workRS = CurrentDb.OpenRecordset(sql)

If workRS.RecordCount = 0 Then
    k = 0.5
Else
    workRS.MoveFirst
    k = workRS![coefficientLightest]
End If
If k = 0 Then k = 0.5

getLightTolerance = k
Exit Function
End Function

Public Sub addToQ(checkResult As String, checkType As String)
Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset

If checkType = "Weight" Or checkType = "Light Weight Check" Then
    If checkResult <> "Pass" Then
        
        'Check si deja enregistre dans la waiting list
        sql = "Select * " & _
                "FROM [waitingList] " & _
                "WHERE [SSCC] ='" & Form_Main.SSCCNumber & "'" & _
                "AND [checkType] = '" & checkType & "';"
        Set workRS = CurrentDb.OpenRecordset(sql)
        If workRS.RecordCount > 0 Then
            result = MsgBox("This handling unit is already waiting for Manual Check")
        End If
        
        'compte nombre de HUlines pour le sscc
        sql2 = "SELECT * FROM [handlingUnitDetails] WHERE [SSCC] = '" & Form_Main.SSCCNumber & "';"
        Set workRS2 = CurrentDb.OpenRecordset(sql2)
        If workRS2.RecordCount > 0 Then
        workRS2.MoveLast
        countLines = workRS2.RecordCount
        End If
        
        'add la sscc to waiting list
        Set Data = CurrentDb
        Set waitingList = Data.OpenRecordset("waitingList")
            waitingList.AddNew
            waitingList("checkDate").Value = Date
            waitingList("checkTime").Value = Time
            waitingList("SSCC").Value = Form_Main.SSCCNumber
            waitingList("Workstation").Value = getWorkstation
            waitingList("checkType").Value = checkType
            waitingList("Result").Value = checkResult
            waitingList("HUlines").Value = countLines
            waitingList.Update
    End If
End If
End Sub

Public Sub checkingSequence(autoManual As String)

On Error GoTo Routine_Error

pickerCheck2 = Form_Main.picker100check.Value
sensitiveMat = Form_Main.sensitiveCheck.Value
testWeight = False
testLight = False

'1 : If Weightcheck checked -> Require to capture the weight -> load and perform scale capture (auto or manually)
If autoManual = "auto" Then
    'from ssccnumber_afterupdate : load scale, auto input
    If Form_Main.weightCheck.Value = -1 And (Form_Main.measuredWeight = 0 Or IsNull(Form_Main.measuredWeight) = True) Then
        DoCmd.OpenForm "weightInput", acNormal
        Form_weightInput.Move 6000, 4000
         Call closeSerialPorts
         Call initSerialPorts
         Call ProcessDataFlow
         Call closeSerialPorts
        DoCmd.Close acForm, "weightInput"
    End If
Else
'from measuredWeight_afterupdate : manual input
End If

'2 : check if HU was picked by a 100% Check picker : if yes -> go to manual check
If Form_Main.pickerCheck.Value = -1 And pickerCheck2 = -1 Then
    displayOrangeLight
    displayWCResults "Picker", 0
    recordCheckResult "Picker", "Pass", 0, Form_Main.measuredWeight.Value, "", "", ""
    addToQ "Pass", "Picker"
    Exit Sub
    
'3 : check if HU contains sensitive material : if yes->go to manual check
ElseIf Form_Main.sensitiveProductCheck = -1 And sensitiveMat = -1 Then
    displayOrangeLight
    displayWCResults "Sensitive", 0
    recordCheckResult "Sensitive", "Pass", 0, Form_Main.measuredWeight.Value, "", "", ""
    addToQ "Pass", "Sensitive"
    Exit Sub
End If

'4 : Check light weight, then check regular weight
If Not weightComparison Then
    displayRedLight
    Exit Sub
Else
    displayGreenLight
    printCheckLabel ("Weight")
    Form_Main.SSCCNumber.SetFocus
    initiateAndon
    cleaningWC
    Exit Sub
End If

'7 : fix abortion or error on scale weightcapture
Routine_Exit:
    Exit Sub

Routine_Error:
Call closeSerialPorts
    Resume Routine_Exit
End
End Sub
