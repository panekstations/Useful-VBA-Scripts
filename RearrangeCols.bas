Sub RearrangeCols()

Dim search As Range
Dim cnt As Integer
Dim colOrdr As Variant
Dim indx As Integer


colOrdr = Application.InputBox("Select the headers you want this tab to keep:", "Replicate Data Columns", Type:=64)


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
