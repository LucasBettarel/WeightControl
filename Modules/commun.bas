Attribute VB_Name = "commun"
'PickPack Quality check
'Tool Designed and developped for Hub Asia by:
'Antoine NICOLE
'Stephen HOUSSAYE
'Lucas BETTAREL


Option Compare Database

Function findAttachment(title As String)

'Requires reference to Microsoft Office 10.0 Object Library.
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Dim FileList As Variant
   FileList = ""
   
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   With fDialog
     
      .AllowMultiSelect = False
      .title = title
      .Filters.Clear
      .Filters.Add "All Files", "*.*"

      If .Show = True Then
     
         For Each varFile In .SelectedItems
        
            GetFileName = varFile
            findAttachment = varFile
               
         Next
      Else
         
      End If
      
   End With

End Function
