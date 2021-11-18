Public Function Rebates(AUM, RateBps, Days)
Rebates = (AUM * (RateBps / 10000) * (Days / 365))
End Function

Public Function RebatesBps(AUM, Rebates, Days)
RebatesBps = ((Rebates / AUM) * 10000 * (365 / Days))
End Function


Public Function ROWCOLMATCH(RowItem, ColItem, rngTable As Range)
    
'like VLOOKUP but you provide the name of the column and row
    
    Dim lngRowMatch       As Long
    Dim lngColMatch       As Long
    On Error GoTo err_handle
    With rngTable
        lngRowMatch = Application.WorksheetFunction.Match(RowItem, .Columns(1), 0)
        lngColMatch = Application.WorksheetFunction.Match(ColItem, .Rows(1), 0)
        ROWCOLMATCH = .Cells(lngRowMatch, lngColMatch)
    End With
    Exit Function
     
     
err_handle:
    ROWCOLMATCH = CVErr(xlErrRef)
End Function

Public Function ISIN(Text As String) As String

'extracts an ISIN anywhere in a cell


FirstNbr = InStr(1, Text, "LU0") + _
            InStr(1, Text, "LU1") + _
            InStr(1, Text, "LU2") + _
            InStr(1, Text, "FR0") + _
            InStr(1, Text, "GB0") + _
            InStr(1, Text, "IE0")

ISIN = Mid(Text, FirstNbr, 12)


End Function


'will extract email address anywhere in a string

Function ExtractEmail(s As String) As String
    Dim AtSignLocation As Long
    Dim i As Long
    Dim TempStr As String
    Const CharList As String = "[A-Za-z0-9._-]"
    
    'Get location of the @
    AtSignLocation = InStr(s, "@")
    If AtSignLocation = 0 Then
        ExtractEmail = "" 'not found
    Else
        TempStr = ""
        'Get 1st half of email address
        For i = AtSignLocation - 1 To 1 Step -1
            If Mid(s, i, 1) Like CharList Then
                TempStr = Mid(s, i, 1) & TempStr
            Else
                Exit For
            End If
        Next i
        If TempStr = "" Then Exit Function
        'get 2nd half
        TempStr = TempStr & "@"
        For i = AtSignLocation + 1 To Len(s)
            If Mid(s, i, 1) Like CharList Then
                TempStr = TempStr & Mid(s, i, 1)
            Else
                Exit For
            End If
        Next i
    End If
    'Remove trailing period if it exists
    If Right(TempStr, 1) = "." Then TempStr = _
       Left(TempStr, Len(TempStr) - 1)
    ExtractEmail = TempStr
End Function


