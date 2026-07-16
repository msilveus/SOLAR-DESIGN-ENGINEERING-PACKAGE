Attribute VB_Name = "modDataEngine"
Public Sub TestPanelToRust()
    Dim PanelID As String
    Dim vmp As Double
    Dim voc As Double
    Dim watts As Double

    PanelID = "P000001"

    vmp = CDbl(GetDBValue("DB_Panels", PanelID, "Vmp V"))
    voc = CDbl(GetDBValue("DB_Panels", PanelID, "Voc V"))
    watts = CDbl(GetDBValue("DB_Panels", PanelID, "STC W"))

    MsgBox _
        "String Vmp: " & EngineStringVmp(vmp, 6) & vbCrLf & _
        "String Voc: " & EngineStringVoc(voc, 6) & vbCrLf & _
        "Array kW: " & EngineArrayDcPower(watts, 36) / 1000
End Sub

Public Function GetColumnIndex(ByVal SheetName As String, ByVal HeaderName As String) As Long
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim c As Long

    Set ws = ThisWorkbook.Worksheets(SheetName)
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).value)) = HeaderName Then
            GetColumnIndex = c
            Exit Function
        End If
    Next c

    Err.Raise vbObjectError + 2001, "GetColumnIndex", _
        "Header not found: " & HeaderName & " in " & SheetName
End Function

Public Function GetRowByID(ByVal SheetName As String, ByVal id As String) As Long
    Dim ws As Worksheet
    Dim idCol As Long
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SheetName)
    idCol = GetColumnIndex(SheetName, "ID")
    lastRow = ws.Cells(ws.Rows.count, idCol).End(xlUp).Row

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, idCol).value)) = id Then
            GetRowByID = r
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 2002, "GetRowByID", _
        "ID not found: " & id & " in " & SheetName
End Function

Public Function GetDBValue(ByVal SheetName As String, ByVal id As String, ByVal FieldName As String) As Variant
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long

    Set ws = ThisWorkbook.Worksheets(SheetName)
    r = GetRowByID(SheetName, id)
    c = GetColumnIndex(SheetName, FieldName)

    GetDBValue = ws.Cells(r, c).value
End Function

