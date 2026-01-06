

Private Sub Filter1_Click()
    Dim ws As Worksheet
    Set ws = Me   'this is Sheet5

    Dim headerRow As Long
    headerRow = 5   'row where your column titles are

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    Else
        ws.Range("A" & headerRow).CurrentRegion.AutoFilter
    End If
End Sub
