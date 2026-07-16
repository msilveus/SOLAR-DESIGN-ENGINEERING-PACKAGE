Attribute VB_Name = "modDropdowns"
Option Explicit

Public Function BuildDropdownCache( _
        ByVal RawList As String, _
        ByVal idColumn As Long, _
        ByVal DisplayColumn As Long, _
        ByVal IDHeader As String, _
        ByVal DisplayHeader As String, _
        ByVal ListName As String)
        
    Dim aValues As Variant
    Dim aItem As Variant
    Dim rowCount As Long
    Dim item As Variant
    Dim ws As Worksheet
    Set ws = DBListsSheet()
    
    ' Clear previous cache (keep the rest of DB_Lists intact)
    ws.Columns(idColumn).ClearContents
    ws.Columns(DisplayColumn).ClearContents
    
    ws.Cells(DBLIST_ROW_HEADER, idColumn).value = IDHeader
    ws.Cells(DBLIST_ROW_HEADER, DisplayColumn).value = DisplayHeader
    
    aValues = Split(RawList, vbLf) ' DLL response

    rowCount = DBLIST_ROW_FIRST_DATA

    For Each item In aValues

        If Len(item) > 0 Then

            aItem = Split(item, vbTab)

            ' Column AA (Panel ID)
            ws.Cells(rowCount, idColumn).value = aItem(0)

            ' Column AB (Display text)
            ws.Cells(rowCount, DisplayColumn).value = aItem(1)

            rowCount = rowCount + 1

        End If

    Next item
    
    SetOrReplaceName ListName, _
    "=" & WS_DBLISTS & "!" & _
    ws.Range( _
        ws.Cells(DBLIST_ROW_FIRST_DATA, DisplayColumn), _
        ws.Cells(rowCount - 1, DisplayColumn)).Address

BuildDropdownCache = rowCount - 1
        
End Function

Public Sub BuildEquipmentDropdowns()
    BuildDropdownCache PanelModels(), _
                       COL_PANEL_ID, _
                       COL_PANEL_DISPLAY, _
                       "Panel ID", _
                       "Panel Model", _
                       NAME_PANEL_LIST
    
    BuildDropdownCache InverterModels(), _
                       COL_INVERTER_ID, _
                       COL_INVERTER_DISPLAY, _
                       "Inverter ID", _
                       "Inverter Model", _
                       NAME_INVERTER_LIST
    
    BuildDropdownCache BatteryModels(), _
                       COL_BATTERY_ID, _
                       COL_BATTERY_DISPLAY, _
                       "Battery ID", _
                       "Battery Model", _
                       NAME_BATTERY_LIST
End Sub

Public Sub ApplyEquipmentDropdowns()

    ApplyListValidation DesignPanelCell(), "=" & NAME_PANEL_LIST
    ApplyListValidation DesignInverterCell(), "=" & NAME_INVERTER_LIST
    ApplyListValidation DesignBatteryCell(), "=" & NAME_BATTERY_LIST

End Sub
Public Sub ApplyBuildOptionDropdowns()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("02_Build")

    ApplyListValidation DesignOrientationCell(), "=" & NAME_ORIENTATION_LIST  ' Orientation
    ApplyListValidation DesignMountTypeCell(), "=" & NAME_MOUNT_TYPE_LIST     ' Mount Type
    ApplyListValidation DesignMpptModeCell(), "=" & NAME_MPPT_MODE_LIST       ' MPPT Mode

End Sub

Private Sub ApplyListValidation(ByVal Target As Range, ByVal formula As String)

    With Target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

End Sub

Public Sub SaveSelectedID( _
        ByVal selectedDisplay As String, _
        ByVal displayName As String, _
        ByVal idColumn As Long, _
        ByVal selectedIDName As String)

    Dim displayRange As Range
    Dim matchCell As Range
    Dim selectedID As String

    Set displayRange = ThisWorkbook.Names(displayName).RefersToRange

    Set matchCell = displayRange.Find( _
        What:=selectedDisplay, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If matchCell Is Nothing Then
        Err.Raise vbObjectError + 3001, "SaveSelectedID", _
            "Selected dropdown value not found: " & selectedDisplay
    End If

    selectedID = DBListsSheet.Cells(matchCell.Row, idColumn).value

    ThisWorkbook.Names(selectedIDName).RefersToRange.value = selectedID

End Sub
