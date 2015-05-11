VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_reportMDissues1"
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

Private Sub Command103812_Click()
DoCmd.Close acForm, "reportMDissues1"
End Sub

Private Sub Command9_Click()

Dim sql3 As String
Dim sql As String

result = MsgBox("Are you sure you entered the right data?", vbYesNo)

If result = vbNo Then
    Exit Sub
Else

DoCmd.SetWarnings False


sql = "DELETE [Local Material Master].* FROM [Local Material Master] " & _
    "INNER JOIN [MDissuesList]" & _
        "ON [Local Material Master].[Material] = [MDissuesList].[material] " & _
        "WHERE [MDissuesList].[removeCheck] = 0 ;"
        
DoCmd.RunSQL sql

sql3 = "INSERT INTO [Local Material Master] ([Material],[Quantity1],[Weight1],[Description],[EAN14],[Quantity2],[EAN13],[Weight2])" & _
        "SELECT [Material],[Unit quantity],[Unit Weight],[description],[EAN14],[Package quantity],[EAN13],[Package Weight] " & _
        "FROM [MDissuesList]  " & _
        "WHERE [removeCheck] = 0 ;"
        
 '"[Local Material Master].[Record Date] = " & Date & ", " & _

DoCmd.RunSQL sql3

DoCmd.SetWarnings True

DoCmd.Close acForm, "reportMDissues1"

End If

Dim theoWeight As Double
Dim actualWeight As Double
theoWeight = 0
actualWeight = 0
recordCheckResult "Manual", "Pass", theoWeight, actualWeight, "", 0, ""

MsgBox "Check result as been recorded as Manual - Pass"

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

'End If

Form_Main.SSCCNumber.SetFocus
initiateAndon
cleaningWC

End Sub

