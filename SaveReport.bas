Sub Savereport()

Dim FolderName As String
With Application.FileDialog(msoFileDialogFolderPicker)
  .AllowMultiSelect = False
  .Show
  On Error Resume Next
  FolderName = .SelectedItems(1)
  Err.Clear
  On Error GoTo 0
End With

    Sheets(ActiveSheet.Name).Select
    Sheets(ActiveSheet.Name).Move
    ActiveWorkbook.SaveAs Filename:=FolderName & "\" & Range("A3") & " " & Range("B3") & ".xlsx", FileFormat:=51
