Attribute VB_Name = "ImportCsvOnOpen9"
Option Explicit

Private Sub ImportCsvToSheet(ByVal csvPath As String, ByVal targetSheetName As String)
    Dim ws As Worksheet
    Dim qt As QueryTable

    If Dir(csvPath) = "" Then
        MsgBox "CSV file not found: " & csvPath, vbExclamation, "SDEP CSV Import"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = targetSheetName
    End If

    ws.Cells.Clear

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With
End Sub

Sub DeleteNameIfExists(ByVal nameText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
End Sub

Function lastRow(ByVal SheetName As String, ByVal colLetter As String) As Long
    With ThisWorkbook.Worksheets(SheetName)
        lastRow = .Cells(.Rows.Count, colLetter).End(xlUp).Row
    End With
End Function
