Attribute VB_Name = "Services_Main_KPI"
'Weight Control
'Tool Designed and developped for Hub Asia by:
'Lucas BETTAREL

Option Compare Database

Public Sub ResGraphCalcul()

Dim grphChart As Object
Dim ChartObjectName As String
Dim lngType As Long, ProductivityType As Long, BarType As Long, FamilyProductivityType As Long


Dim sql As String
Dim operation As String
Dim timeunit As String
Dim word As String
Dim word2 As String
Dim word3 As String
Dim word4 As String
Dim word5 As String
Dim MonthNumb As String
Dim RS As Recordset

''''Filtre
'Picking Quality PPM";"Rejection Rate"

If Form_Main.KPI = "" Then Exit Sub
If IsNull(Form_Main.KPI) = True Then Exit Sub

If Form_Main.KPI = "Rejection Rate" Then

'work: select column
'word5 : name to display on legend
'word3: condition

    word = "ROUND((SUM ([weightFailHU]) / IIF((SUM([weightPassHU]) + SUM([weightFailHU])) = 0,1,(SUM([weightPassHU]) + SUM([weightFailHU]))))*100,0) as WeightRejectionPercentage,ROUND((SUM([EANFailHU]) / IIF((SUM([EANPassHU]) + SUM([EANFailHU]))=0,1,(SUM([EANPassHU]) + SUM([EANFailHU])))) * 100, 0) As EANvisualRejectionPercentage"
    word5 = ""
    word3 = ""
    word3 = ""
  BarType = xlColumnClustered
End If

If Form_Main.KPI = "Picking Quality PPM" Then
'totalLinesPassed
'totalLinesFailed
    word = "ROUND((SUM ([totalLinesFailed]) / IIF(SUM([totalLinesPassed])=0,1,SUM([totalLinesPassed])))* 1000000,0)  as PPM "
    word5 = ""
    word3 = ""
BarType = xlLine
End If


If Form_Main.KPI = "Weight Details" Then

    word = "ROUND(SUM ([weightFailHU]),0) as HU_weightFail , ROUND(SUM ([weightPassHU]),0) as HU_weightPass "
    word5 = ""
    word3 = ""
    word3 = ""
  BarType = xlColumnStacked
End If

If Form_Main.KPI = "EAN/Visual Details" Then

    word = "ROUND(SUM ([EANFailHU]),0) as HU_EANFail , ROUND(SUM ([EANPassHU]),0) as HU_EANPass "
    word5 = ""
    word3 = ""
    word3 = ""
  BarType = xlColumnStacked
End If

If Form_Main.ResTimeUnit = "Month" Then timeunit = "Month([checkDate])"
If Form_Main.ResTimeUnit = "Day" Then timeunit = "[checkDate]"
If Form_Main.ResTimeUnit = "Week" Then timeunit = "format([checkDate],'ww')"


If Form_Main.MonthLenght = "" Or IsNull(Form_Main.MonthLenght) = True Then

    MonthNumb = "Month([checkDate])Like '*' "
Else

    MonthNumb = "Month([checkDate])='" & Form_Main.MonthLenght & "'"

End If
               
    sql = "SELECT " & timeunit & " AS TimeLine, " & word & word5 & _
     " FROM [KPI]" & _
     " WHERE Year([checkDate])= '" & Form_Main.ResTimeLenght & "'" & word3 & "" & _
     " And " & MonthNumb & "" & _
     " GROUP BY " & timeunit & ";"
       
Set RS = CurrentDb.OpenRecordset(sql)


If RS.RecordCount = 0 Then
    Form_Main.NoData.Visible = True
Else
    Form_Main.NoData.Visible = False
End If

Form_Main.GraphResults.RowSourceType = "Table/Query"
                
Form_Main.GraphResults.RowSource = sql

'''''''''''''''''''''''''''''''
'''''' Change Form of the graph
'''''''''''''''''''''''''''''''


ChartObjectName = "GraphResults"

'ProductivityType = xlLineMarkersStacked
'FamilyProductivityType = xlLine

'BarType = xlColumnStacked
'BarType = xlColumnClustered
    lngType = BarType

''' Different types of graphs

'xlColumnClustered , xlColumnStacked, xlColumnStacked100 ''Column
'xlBarClustered , xlBarStacked, xlBarStacked100 ''Bar
'xlLine , xlLineMarkersStacked, xlLineStacked '' Line
'xlPie , xlPieOfPie ''Pie
'xlXYScatter , xlXYScatterLines ''Scatter

Form_Main.GraphResults.ChartType = lngType
Form_Main.GraphResults.Requery

End Sub

Public Sub updateKPI()

Dim sql As String
Dim workRS As Recordset
Dim sql2 As String
Dim workRS2 As Recordset
Dim sql3 As String
Dim workRS3 As Recordset

Dim sql4 As String
Dim workRS4 As Recordset

Dim sql5 As String
Dim workRS5 As Recordset
Dim sql6 As String
Dim workRS6 As Recordset

Dim sql7 As String
Dim workRS7 As Recordset

Dim sql8 As String
Dim workRS8 As Recordset

updateFrom = Date

sql = "Select * " & _
        "FROM [KPI] ;"

Set workRS = CurrentDb.OpenRecordset(sql)

If workRS.RecordCount = 0 Then

'update with all data available

sql2 = "Select min([checkDate]) as minDate " & _
        "FROM [QCcheckLog] ;"

Set workRS2 = CurrentDb.OpenRecordset(sql2)
    If workRS2.RecordCount > 0 Then

        updateFrom = workRS2![minDate]

    Else
        MsgBox "No data in QCcheckLog"
        Exit Sub
    End If


Else

workRS.MoveLast

updateFrom = workRS![checkDate]

'update from last date

'End If
End If

'**********update KPI from updateDATE

Do While updateFrom <= Date

Set Data = CurrentDb
Set checkLog = Data.OpenRecordset("KPI")


DoCmd.SetWarnings False
Dim sql9 As String
sql9 = "DELETE * FROM [KPI] WHERE [checkDate] = # " & Format(updateFrom, "mm/dd/yyyy") & " #;"
DoCmd.RunSQL sql9
DoCmd.SetWarnings True


sql3 = "SELECT count([handlingUnit]) AS weightPass FROM (SELECT [handlingUnit]" & _
        "FROM [QCcheckLog] " & _
        "WHERE [checkDate] = #" & Format(updateFrom, "mm/dd/yyyy") & "#" & _
        "AND ([checkType] = 'Weight' OR [checkType] = 'Light Weight Check')" & _
        "AND [Result] = 'Pass' " & _
        "GROUP BY [handlingUnit]);"
       
Set workRS3 = CurrentDb.OpenRecordset(sql3)

'With CurrentDb
' Set qdf = .CreateQueryDef("tmpProductInfo", sql3)
' DoCmd.OpenQuery "tmpProductInfo"
 '.QueryDefs.Delete "tmpProductInfo"
 'End With

sql4 = "SELECT count([handlingUnit]) AS weightFail FROM (SELECT [handlingUnit]" & _
        "FROM [QCcheckLog] " & _
        "WHERE [checkDate] = #" & Format(updateFrom, "mm/dd/yyyy") & "#" & _
        "AND ([checkType] = 'Weight' OR [checkType] = 'Light Weight Check')" & _
        "AND [Result] <> 'Pass' " & _
        "GROUP BY [handlingUnit]);"
       
Set workRS4 = CurrentDb.OpenRecordset(sql4)

sql5 = "SELECT count([handlingUnit]) AS EANpass FROM (SELECT [handlingUnit]" & _
        "FROM [QCcheckLog] " & _
        "WHERE [checkDate] = #" & Format(updateFrom, "mm/dd/yyyy") & "#" & _
        "AND ([checkType] = 'EAN' OR [checkType] = 'Manual')" & _
        "AND [Result] = 'Pass' " & _
        "GROUP BY [handlingUnit]);"
       
Set workRS5 = CurrentDb.OpenRecordset(sql5)

sql6 = "SELECT count([handlingUnit]) AS EANfail FROM (SELECT [handlingUnit]" & _
        "FROM [QCcheckLog] " & _
        "WHERE [checkDate] = #" & Format(updateFrom, "mm/dd/yyyy") & "#" & _
        "AND ([checkType] = 'EAN' OR [checkType] = 'Manual')" & _
        "AND [Result] <> 'Pass' " & _
        "GROUP BY [handlingUnit]);"
       
Set workRS6 = CurrentDb.OpenRecordset(sql6)

sql7 = "SELECT SUM([linesByHU]) AS totalLines FROM (SELECT [handlingUnit],MAX([HUlines])as linesByHU " & _
        "FROM [QCcheckLog] " & _
        "WHERE [checkDate] = #" & Format(updateFrom, "mm/dd/yyyy") & "#" & _
        "GROUP BY [handlingUnit]);"
       
Set workRS7 = CurrentDb.OpenRecordset(sql7)


sql8 = "SELECT count([SSCC]) AS issues " & _
        "FROM [pickingIssuesReport] " & _
        "WHERE [Checking Date] = #" & Format(updateFrom, "mm/dd/yyyy") & "#;"

Set workRS8 = CurrentDb.OpenRecordset(sql8)



If (workRS3.RecordCount > 0 And workRS3![weightPass] > 0) Or (workRS4.RecordCount > 0 And workRS4![WeightFail] > 0) Or (workRS6.RecordCount > 0 And workRS6![EANfail] > 0) Or (workRS5.RecordCount > 0 And workRS5![EANpass] > 0) Then

            checkLog.AddNew
            checkLog("checkDate").Value = updateFrom
            If IsNull(workRS3![weightPass]) = False Then checkLog("weightPassHU").Value = workRS3![weightPass]
            If IsNull(workRS4![WeightFail]) = False Then checkLog("weightFailHU").Value = workRS4![WeightFail]
            If IsNull(workRS5![EANpass]) = False Then checkLog("EANPassHU").Value = workRS5![EANpass]
            If IsNull(workRS6![EANfail]) = False Then checkLog("EANFailHU").Value = workRS6![EANfail]
             If IsNull(workRS7![totalLines]) = False Then checkLog("totalLinesPassed").Value = workRS7![totalLines]
             If IsNull(workRS8![issues]) = False Then checkLog("totalLinesFailed").Value = workRS8![issues]
           
            checkLog.Update

End If
updateFrom = updateFrom + 1
Loop

End Sub


