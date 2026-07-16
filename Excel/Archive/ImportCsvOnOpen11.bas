Attribute VB_Name = "ImportCsvOnOpen11"
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

Sub ApplyBuildDropdowns()

    With Sheets("02_Build").Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=PanelList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With Sheets("02_Build").Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=InverterList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With Sheets("02_Build").Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=BatteryList"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

End Sub

Sub CreateNamedLists()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    DeleteNameIfExists "PanelList"
    DeleteNameIfExists "InverterList"
    DeleteNameIfExists "BatteryList"

    wb.Names.Add Name:="PanelList", _
        RefersTo:="=DB_Panels!$B$2:$B$" & lastRow("DB_Panels", "A")

    wb.Names.Add Name:="InverterList", _
        RefersTo:="=DB_Inverters!$B$2:$B$" & lastRow("DB_Inverters", "A")

    wb.Names.Add Name:="BatteryList", _
        RefersTo:="=DB_Batteries!$B$2:$B$" & lastRow("DB_Batteries", "A")

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
