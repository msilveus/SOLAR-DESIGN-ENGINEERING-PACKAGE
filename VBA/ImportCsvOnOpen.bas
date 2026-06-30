Attribute VB_Name = "ImportCsvOnOpen"
Option Explicit

' SDEP CSV Import Module
' Save the workbook as .xlsm, import this module, then add the Workbook_Open
' call below to ThisWorkbook:
'
' Private Sub Workbook_Open()
'     ImportAllSdepCsv
' End Sub
'
' CSV files are expected in a Data folder next to the workbook:
'   Data\Panels.csv
'   Data\Inverters.csv
'   Data\Batteries.csv
'   Data\Lists.csv

Public Sub ImportAllSdepCsv()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ImportCsvToSheet ThisWorkbook.Path & "\Data\Panels.csv", "DB_Panels"
    ImportCsvToSheet ThisWorkbook.Path & "\Data\Inverters.csv", "DB_Inverters"
    ImportCsvToSheet ThisWorkbook.Path & "\Data\Batteries.csv", "DB_Batteries"
    ImportCsvToSheet ThisWorkbook.Path & "\Data\Lists.csv", "DB_Lists"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

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
