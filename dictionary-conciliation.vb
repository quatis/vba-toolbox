Sub headcountConciliation()
    Dim liveWs As Worksheet, previousWs As Worksheet, resultWs As Worksheet
    Dim previousDict As Object, liveDict As Object, cel As Range
    Dim liveLastLine As Long, previousLastLine As Long, linhaRes As Long
    Dim idCol As Long, colNome As Long, colEntity As Long
    Dim terminationCount As Integer

    Set liveWs = ThisWorkbook.Sheets("FullHC-1")
    Set previousWs = ThisWorkbook.Sheets("FullHC-2")
    
    On Error Resume Next
    Set resultWs = ThisWorkbook.Sheets("Desligamentos")
    On Error GoTo 0
    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Sheets.Add
        resultWs.Name = "Desligamentos"
    End If
    resultWs.Cells.Clear
    
    idCol = indexColumnLookup(previousWs, "ID SENIOR")
    colNome = indexColumnLookup(previousWs, "NAME")
    colEntity = indexColumnLookup(previousWs, "ENTITY")
    
    previousLastLine = previousWs.Cells(previousWs.Rows.Count, idCol).End(xlUp).Row
    liveLastLine = liveWs.Cells(liveWs.Rows.Count, idCol).End(xlUp).Row
    
    Set previousDict = CreateObject("Scripting.Dictionary")
    Set liveDict = CreateObject("Scripting.Dictionary")
    
    For Each cel In previousWs.Range(previousWs.Cells(2, idCol), previousWs.Cells(previousLastLine, idCol))
        If Trim(cel.Offset(0, colEntity - idCol).Value) = "NDBR" Then
            previousDict(Trim(cel.Value)) = Trim(cel.Offset(0, colNome - idCol).Value)
        End If
    Next cel
    
    For Each cel In liveWs.Range(liveWs.Cells(2, idCol), liveWs.Cells(liveLastLine, idCol))
        If Trim(cel.Offset(0, colEntity - idCol).Value) = "NDBR" Then
            liveDict(Trim(cel.Value)) = True
        End If
    Next cel
    
    resultWs.Cells(1, 1).Value = "ID Senior"
    resultWs.Cells(1, 2).Value = "Nome"
    resultWs.Cells(1, 3).Value = "Status"
    linhaRes = 2
    
    Dim chave As Variant
    For Each chave In previousDict.Keys
        If Not liveDict.Exists(chave) Then
            resultWs.Cells(linhaRes, 1).Value = chave
            resultWs.Cells(linhaRes, 2).Value = previousDict(chave)
            resultWs.Cells(linhaRes, 3).Value = "Desligamento"
            linhaRes = linhaRes + 1
            terminationCount = terminationCount + 1
        End If
    Next chave
    
    resultWs.Cells(1, 4).Value = "Total de Desligamentos:"
    resultWs.Cells(1, 5).Value = terminationCount
    resultWs.Columns("A:E").AutoFit
End Sub
