Option Explicit

Public Sub Import_Click()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsOut As Worksheet: Set wsOut = wb.Worksheets("Hedging Adjustment")
    Dim wsMap As Worksheet: Set wsMap = wb.Worksheets("Sheet1")
    
    Dim folderPath As String
    Dim dateLabel As String
    
    folderPath = Trim(CStr(wsOut.Range("O2").Value))
    dateLabel = Trim(CStr(wsOut.Range("O1").Value))
    
    If Len(folderPath) = 0 Then
        MsgBox "Cell O2 (folder path) is blank on Hedging Adjustment.", vbExclamation
        Exit Sub
    End If
    
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Dim mapLastRow As Long
    mapLastRow = wsMap.Cells(wsMap.Rows.Count, "B").End(xlUp).Row
    If mapLastRow < 2 Then
        MsgBox "Sheet1 mapping table looks empty (nothing under column B).", vbExclamation
        GoTo CleanExit
    End If
    
    Dim r As Long
    For r = 2 To mapLastRow
        
        Dim wbBase As String, srcSheetName As String, bookCode As String
        wbBase = Trim(CStr(wsMap.Cells(r, "B").Value))
        srcSheetName = Trim(CStr(wsMap.Cells(r, "C").Value))
        bookCode = Trim(CStr(wsMap.Cells(r, "A").Value))
        
        If Len(wbBase) = 0 Or Len(srcSheetName) = 0 Then
            'skip blank lines
            GoTo NextR
        End If
        
        If Len(bookCode) = 0 Then
            bookCode = InferBookFromName(wbBase)
        End If
        
        'Column letters from Sheet1
        Dim colSec As String, colRef As String, colInc As String, colDeDes As String
        Dim colTransit As String, colSrcCcy As String, colNotional As String, colLtd As String
        
        colSec = Trim(CStr(wsMap.Cells(r, "D").Value))
        colRef = Trim(CStr(wsMap.Cells(r, "E").Value))
        colInc = Trim(CStr(wsMap.Cells(r, "F").Value))
        colDeDes = Trim(CStr(wsMap.Cells(r, "G").Value))
        colTransit = Trim(CStr(wsMap.Cells(r, "H").Value))
        colSrcCcy = Trim(CStr(wsMap.Cells(r, "I").Value))
        colNotional = Trim(CStr(wsMap.Cells(r, "J").Value))
        colLtd = Trim(CStr(wsMap.Cells(r, "K").Value))
        
        If Len(colSec) = 0 Or Len(colLtd) = 0 Then
            Debug.Print "Row " & r & " skipped, missing Security ID or LTD column letter."
            GoTo NextR
        End If
        
        Dim filePath As String
        filePath = ResolveWorkbookPath(folderPath, wbBase, dateLabel)
        
        If Len(filePath) = 0 Then
            Debug.Print "File not found for row " & r & ": " & wbBase
            GoTo NextR
        End If
        
        Dim srcWb As Workbook
        Dim srcWs As Worksheet
        
        On Error GoTo OpenFail
        Set srcWb = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True, Notify:=False, AddToMru:=False)
        On Error GoTo 0
        
        On Error GoTo SheetFail
        Set srcWs = srcWb.Worksheets(srcSheetName)
        On Error GoTo 0
        
        'Find last row in Security ID column (start at row 6)
        Dim startRow As Long: startRow = 6
        Dim srcLast As Long
        srcLast = srcWs.Cells(srcWs.Rows.Count, colSec).End(xlUp).Row
        
        If srcLast < startRow Then
            srcWb.Close SaveChanges:=False
            GoTo NextR
        End If
        
        Dim destStart As Long
        destStart = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row + 1
        If destStart < 2 Then destStart = 2
        
        Dim n As Long
        n = srcLast - startRow + 1
        
        'Write Book column (C)
        wsOut.Range(wsOut.Cells(destStart, "C"), wsOut.Cells(destStart + n - 1, "C")).Value = bookCode
        
        'Copy helper (only copy if column letter provided)
        CopyColValues srcWs, colSec, startRow, srcLast, wsOut, "A", destStart
        CopyColValues srcWs, colRef, startRow, srcLast, wsOut, "B", destStart
        CopyColValues srcWs, colInc, startRow, srcLast, wsOut, "D", destStart
        CopyColValues srcWs, colDeDes, startRow, srcLast, wsOut, "E", destStart
        CopyColValues srcWs, colTransit, startRow, srcLast, wsOut, "F", destStart
        CopyColValues srcWs, colSrcCcy, startRow, srcLast, wsOut, "G", destStart
        CopyColValues srcWs, colNotional, startRow, srcLast, wsOut, "H", destStart
        CopyColValues srcWs, colLtd, startRow, srcLast, wsOut, "I", destStart
        
        srcWb.Close SaveChanges:=False
        
NextR:
        DoEvents
    Next r
    
    MsgBox "Import complete.", vbInformation
    
CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Sub

OpenFail:
    Debug.Print "Could not open: " & filePath
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0
    Resume NextR

SheetFail:
    Debug.Print "Sheet not found in " & filePath & " -> " & srcSheetName
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0
    Resume NextR

End Sub

Private Sub CopyColValues(ByVal srcWs As Worksheet, ByVal colLetter As String, ByVal startRow As Long, ByVal endRow As Long, _
                         ByVal destWs As Worksheet, ByVal destCol As String, ByVal destStartRow As Long)
    If Len(Trim$(colLetter)) = 0 Then Exit Sub
    
    Dim srcRng As Range, destRng As Range
    Set srcRng = srcWs.Range(colLetter & startRow & ":" & colLetter & endRow)
    Set destRng = destWs.Range(destCol & destStartRow).Resize(srcRng.Rows.Count, 1)
    
    destRng.Value = srcRng.Value
End Sub

Private Function ResolveWorkbookPath(ByVal folderPath As String, ByVal wbBase As String, ByVal dateLabel As String) As String
    'Tries to find a matching Excel file in the folder
    Dim f As String
    
    'Most common: base + date somewhere + excel extension
    f = Dir(folderPath & "\" & wbBase & "*" & dateLabel & "*.xl*")
    If Len(f) > 0 Then
        ResolveWorkbookPath = folderPath & "\" & f
        Exit Function
    End If
    
    'Fallback: base without date
    f = Dir(folderPath & "\" & wbBase & "*.xl*")
    If Len(f) > 0 Then
        ResolveWorkbookPath = folderPath & "\" & f
        Exit Function
    End If
    
    ResolveWorkbookPath = ""
End Function

Private Function InferBookFromName(ByVal s As String) As String
    'Pulls leading letters from workbook base name (ex: "CICIT11600..." -> "CICIT")
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            out = out & ch
        Else
            Exit For
        End If
    Next i
    InferBookFromName = out
End Function
