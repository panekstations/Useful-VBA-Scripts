Sub Savefile()

'checks to make sure chees isn't empty
If WorksheetFunction.CountA(Range("A1:Z100")) = 0 Then
    MsgBox "This appears to be an empty sheet."

Else
    
    Dim FolderName As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
    On Error Resume Next
        FolderName = .SelectedItems(1)
        Err.Clear
    On Error GoTo 0
End With

'determines if it's the only sheet in the workbook
If Application.Sheets.Count <> 1 Then

    ActiveWindow.SelectedSheets.Select
    ActiveWindow.SelectedSheets.Move
    
End If
    
    'If saving file that it not normal calculation
    
    If WorksheetFunction.CountA(Range("A3:C3")) = 0 Then
        Name = Application.InputBox("Cells A3-C3 are blank. Enter a filename")
        ActiveWorkbook.SaveAs Filename:=FolderName & "\" & Name & ".xlsx", FileFormat:=51
    Else
        ActiveWorkbook.SaveAs Filename:=FolderName & "\" & Range("A3") & " " & Range("C3") & ".xlsx", FileFormat:=51
    End If
    
End If
