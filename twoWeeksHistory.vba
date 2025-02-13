Sub saveHistory()
    Dim originWs As Worksheet, newWs As Worksheet
    Dim oldWs As Worksheet
    Dim newWsName As String, oldWsName As String, deletionWsName As String
    Dim today As Date
    Dim lastCol As Long
    Dim tbl As ListObject

    Application.ScreenUpdating = False
    Set originWs = ThisWorkbook.Sheets("Posições EC")
    today = Date
    newWsName = Format(today - 7, "DD.MM")
    oldWsName = Format(today - 14, "DD.MM")
    deletionWsName = Format(today - 21, "DD.MM")

    On Error Resume Next
    Set oldWs = ThisWorkbook.Sheets(deletionWsName)
    If Not oldWs Is Nothing Then
        Application.DisplayAlerts = False
        oldWs.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = newWsName
    originWs.Cells.Copy
    newWs.Cells.PasteSpecial Paste:=xlPasteValues

    lastCol = newWs.Cells(1, Columns.Count).End(xlToLeft).Column
    lastRow = newWs.Cells(Rows.Count, 1).End(xlUp).Row
    Set tbl = newWs.ListObjects.Add(xlSrcRange, newWs.Range(newWs.Cells(1, 1), newWs.Cells(lastRow, lastCol)), , xlYes)
    tbl.Name = "Tabela_" & Replace(newWsName, ".", "_")

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

