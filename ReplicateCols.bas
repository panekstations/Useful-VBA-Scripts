Sub ColumnsToKeep()


'Columns cleaning

Dim rng As Range
Set rng = Application.InputBox("Select the headers you want this tab to keep:", "Replicate Data Columns", Type:=8)


    Mylist = rng
    LC = Cells(1, Columns.Count).End(xlToLeft).Column
    
    For mycol = LC To 1 Step -1
        x = ""
        On Error Resume Next
        x = WorksheetFunction.Match(Cells(1, mycol), Mylist, 0)
        If Not IsNumeric(x) Then Columns(mycol).EntireColumn.Delete
    Next mycol
        Erase arrColumnNames


'rearrange columns by Header
    Dim search As Range
    Dim cnt As Integer
    Dim colOrdr As Variant
    Dim indx As Integer
    
    colOrdr = Array(rng)  'define column order using the array aboved
    
    cnt = 1
    
    
    For indx = LBound(colOrdr) To UBound(colOrdr)
        Set search = Rows("1:1").Find(colOrdr(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                Columns(cnt).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        cnt = cnt + 1
        End If
    Next indx


End Sub
