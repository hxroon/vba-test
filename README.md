Option Explicit

'========================
' Helpers
'========================

'fast value copy, no clipboard, no 1004
Public Sub CopyValuesCol( _
    ByVal srcWS As Worksheet, ByVal srcColLetter As String, ByVal srcStartRow As Long, ByVal srcLastRow As Long, _
    ByVal dstWS As Worksheet, ByVal dstCol As Long, ByVal dstStartRow As Long)

    Dim n As Long
    If Len(srcColLetter) = 0 Then Exit Sub
    n = srcLastRow - srcStartRow + 1
    If n <= 0 Then Exit Sub

    dstWS.Cells(dstStartRow, dstCol).Resize(n, 1).Value = _
        srcWS.Range(UCase$(srcColLetter) & srcStartRow).Resize(n, 1).Value
End Sub

'find destination column by header text in row 1
Private Function ColByHeader(ws As Worksheet, headerText As String) As Long
    Dim m
    m = Application.Match(headerText, ws.Rows(1), 0)
    If IsError(m) Then
        ColByHeader = 0
    Else
        ColByHeader = CLng(m)
    End If
End Function

'last used row in a given column
Private Function LastRowIn(ws As Worksheet, col As Long) As Long
    If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        LastRowIn = 1
    Else
        LastRowIn = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    End If
End Function

'derive a book code like "CMA" or "MUNIPTAX" from the workbook name
Private Function BookCodeFromName(ByVal bookName As String) As String
    Dim s As String, i As Long, ch As String
    s = Trim(bookName)
    'take leading letters until a digit or space
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            BookCodeFromName = BookCodeFromName & ch
        Else
            Exit For
        End If
    Next i
    If Len(BookCodeFromName) = 0 Then BookCodeFromName = s
End Function

'========================
' Main importer that uses Sheet1 mapping
'========================
Public Sub Import_DEV()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDst As Worksheet: Set wsDst = wb.Worksheets("Hedging Adjustment")
    Dim wsMap As Worksheet: Set wsMap = wb.Worksheets("Sheet1")

    Dim fp As String, d1 As String
    fp = Trim(wsDst.Range("O2").Value)        'folder path
    d1 = Trim(wsDst.Range("Q1").Value)        'month token, eg "Sep2025"
    If Len(fp) = 0 Then
        MsgBox "Put the local folder path in Hedging Adjustment!O2", vbExclamation
        Exit Sub
    End If
    If Right$(fp, 1) <> "\" And Right$(fp, 1) <> "/" Then fp = fp & "\"

    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    'identify mapping header columns on Sheet1, from column D to last
    Dim firstMapCol As Long: firstMapCol = 4
    Dim lastMapCol As Long: lastMapCol = wsMap.Cells(1, wsMap.Columns.Count).End(xlToLeft).Column
    If lastMapCol < firstMapCol Then
        MsgBox "No mapping headers found on Sheet1 row 1", vbExclamation
        GoTo CleanExit
    End If

    'walk each workbook row on Sheet1, starting row 3
    Dim r As Long, lastMapRow As Long
    lastMapRow = LastRowIn(wsMap, 2)
    If lastMapRow < 3 Then
        MsgBox "No workbook names found on Sheet1 column B", vbExclamation
        GoTo CleanExit
    End If

    Dim bookName As String, sheetName As String
    Dim srcWB As Workbook, srcWS As Worksheet
    Dim openPath As String
    Dim LRdst As Long, LR2 As Long
    Dim srcLast As Long
    Dim c As Long, dstCol As Long
    Dim fieldName As String
    Dim bookCode As String

    For r = 3 To lastMapRow
        bookName = Trim(wsMap.Cells(r, 2).Value)  'col B
        sheetName = Trim(wsMap.Cells(r, 3).Value) 'col C
        If Len(bookName) = 0 Then GoTo NextRow

        bookCode = BookCodeFromName(bookName)
        openPath = fp & bookName & " - " & d1

        'open safely
        Set srcWB = Nothing
        On Error Resume Next
        Set srcWB = Workbooks.Open(openPath, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0
        If srcWB Is Nothing Then
            Debug.Print "Missing file: " & openPath
            GoTo NextRow
        End If

        'get the sheet
        If Len(sheetName) = 0 Then sheetName = "Financial Reporting Summary"
        Set srcWS = Nothing
        On Error Resume Next
        Set srcWS = srcWB.Worksheets(sheetName)
        On Error GoTo 0
        If srcWS Is Nothing Then
            Debug.Print "Missing sheet '" & sheetName & "' in: " & openPath
            srcWB.Close False
            GoTo NextRow
        End If

        'compute source last row from column A
        srcLast = LastRowIn(srcWS, 1)
        If srcLast < 6 Then
            srcWB.Close False
            GoTo NextRow
        End If

        'destination start row
        LRdst = LastRowIn(wsDst, 1) + 1

        'copy each mapped field according to headers on row 1 of Sheet1
        For c = firstMapCol To lastMapCol
            fieldName = Trim(wsMap.Cells(1, c).Value)       'destination header name
            If Len(fieldName) = 0 Then GoTo NextField

            dstCol = ColByHeader(wsDst, fieldName)          'find the destination column by header on Hedging Adjustment
            If dstCol = 0 Then
                'header not found on destination, skip silently
                GoTo NextField
            End If

            'source column letter for this field, per current row
            Dim srcColLetter As String
            srcColLetter = Trim(wsMap.Cells(r, c).Value)     'eg "A", "B", "F", "AH"
            If Len(srcColLetter) > 0 Then
                CopyValuesCol srcWS, srcColLetter, 6, srcLast, wsDst, dstCol, LRdst
            End If

NextField:
        Next c

        'fill Column C with book code for the newly pasted rows
        LR2 = LastRowIn(wsDst, 1)
        wsDst.Range(wsDst.Cells(LRdst, 3), wsDst.Cells(LR2, 3)).Value = bookCode

        'close source
        srcWB.Close SaveChanges:=False
        Set srcWB = Nothing
        Set srcWS = Nothing

NextRow:
    Next r

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Exit Sub

CleanFail:
    'basic fail-safe to restore app state
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Err.Raise Err.Number, "Import_DEV", Err.Description
End Sub

'optional, quick cleanup similar to your Filter button
Public Sub Filter_DEV()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Hedging Adjustment")
    'turn on filter on header row 1
    With ws
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        .Rows(1).AutoFilter
    End With
End Sub
