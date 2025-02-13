
Function checkEntity(ByVal ws As Worksheet, ByVal nome As String) As Boolean
    Dim cel As Range, lastLine As Long
    Dim nameCol As Long, entityCol As Long
    
    nameCol = indexColumnLookup(ws, "NAME")
    entityCol = indexColumnLookup(ws, "ENTITY")
    lastLine = ws.Cells(ws.Rows.Count, colNome).End(xlUp).Row
    
    For Each cel In ws.Range(ws.Cells(2, nameCol), ws.Cells(lastLine, nameCol))
        If Trim(cel.Value) = Trim(nome) And Trim(cel.Offset(0, entityCol - colNome).Value) = "NDBR" Then
            isEntity = True
            Exit Function
        End If
    Next cel
    
    isEntity = False
End Function

Function indexColumnLookup(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(1, i).Value) = Trim(headerName) Then
            indexColumnLookup = i
            Exit Function
        End If
    Next i
    MsgBox "Coluna " & headerName & " não encontrada!", vbCritical
    indexColumnLookup = 0
End Function
