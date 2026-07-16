Attribute VB_Name = "modWorkbookLayout"
Option Explicit

Public Const PROTECT_WORKBOOK As Boolean = True

'==============================================================
' Worksheet Status
'==============================================================
Public Const STATUS_PASS As String = "PASS"
Public Const STATUS_WARN As String = "WARN"
Public Const STATUS_FAIL As String = "FAIL"

'==============================================================
' Design Status
'==============================================================
Public Enum DesignStatus
    DS_PASS = 1
    DS_WARN = 2
    DS_FAIL = 3
End Enum

'==============================================================
' Worksheet Names
'==============================================================

Public Const WS_DASHBOARD As String = "01_Dashboard"
Public Const WS_BUILD As String = "02_Build"
Public Const WS_LAYOUT As String = "03_Layout_Designer"
Public Const WS_SHADE As String = "04_Shade_Optimizer"
Public Const WS_REVIEW As String = "05_Design_Review"
Public Const WS_DBLISTS As String = "DB_Lists"

'==============================================================
' Build Sheet Static Option Dropdown Cells
'==============================================================

Public Const BUILD_ORIENTATION_CELL As String = "B9"
Public Const BUILD_MOUNT_TYPE_CELL As String = "B10"
Public Const BUILD_MPPT_MODE_CELL As String = "B20"

'==============================================================
' DB_Lists Static List Columns
' Existing manually maintained lists live on the left side.
'==============================================================

Public Const DBLIST_ROW_HEADER As Long = 1
Public Const DBLIST_ROW_FIRST_DATA As Long = 2

Public Const COL_ORIENTATION_LIST As Long = 1       'A
Public Const COL_MOUNT_TYPE_LIST As Long = 2        'B
Public Const COL_MPPT_MODE_LIST As Long = 3         'C

'==============================================================
' DB_Lists Runtime Equipment Cache Columns
' Runtime-generated data starts at AA and is owned by VBA.
'==============================================================

Public Const COL_PANEL_ID As Long = 27              'AA
Public Const COL_PANEL_DISPLAY As Long = 28         'AB

Public Const COL_INVERTER_ID As Long = 30           'AD
Public Const COL_INVERTER_DISPLAY As Long = 31      'AE

Public Const COL_BATTERY_ID As Long = 33            'AG
Public Const COL_BATTERY_DISPLAY As Long = 34       'AH

'==============================================================
' DB_Lists Hidden Selected-ID State
'==============================================================

Public Const COL_SELECTED_ID_LABEL As Long = 40     'AN
Public Const COL_SELECTED_ID_VALUE As Long = 41     'AO

Public Const ROW_SELECTED_PANEL As Long = 2
Public Const ROW_SELECTED_INVERTER As Long = 3
Public Const ROW_SELECTED_BATTERY As Long = 4

'==============================================================
' Build Design Cells
'==============================================================
Public Const BUILD_INPUT_FIRST_ROW As Long = 4
Public Const BUILD_INPUT_LAST_ROW As Long = 23
Public Const COL_DESIGN_INPUTS As Long = 2
Public Const BUILD_PANEL_CELL As Long = 5
Public Const BUILD_INVERTER_CELL As Long = 6
Public Const BUILD_BATTERY_CELL As Long = 7
Public Const BUILD_BATTERY_QTY As Long = 8
Public Const BUILD_ORIENTATION As Long = 9
Public Const BUILD_MOUNT_TYPE As Long = 10
Public Const BUILD_BANK_HEIGHT As Long = 11
Public Const BUILD_BANK_WIDTH As Long = 12
Public Const BUILD_BANK_COUNT As Long = 13
Public Const BUILD_BANK_TILT As Long = 14
Public Const BUILD_BANK_AZIMUTH As Long = 15
Public Const BUILD_DESIGN_MIN_TEMP_C As Long = 16
Public Const BUILD_DESIGN_MAX_TEMP_C As Long = 17
Public Const BUILD_PANELS_PER_STRING As Long = 18
Public Const BUILD_TOTAL_STRINGS As Long = 19
Public Const BUILD_STRINGS_PER_MPPT As Long = 20
Public Const BUILD_MPPT_MODE As Long = 21
Public Const BUILD_BOS_ESTIMATE As Long = 22

'==============================================================
' Build Equipment Cells
'==============================================================
Public Const COL_EQUIPMENT_VALUES As Long = 4
Public Const EQUIPMENT_PANEL_WATTS As Long = 4
Public Const EQUIPMENT_PANEL_VOC As Long = 5
Public Const EQUIPMENT_PANEL_VMP As Long = 6
Public Const EQUIPMENT_PANEL_ISC As Long = 7
Public Const EQUIPMENT_PANEL_IMP As Long = 8
Public Const EQUIPMENT_PANEL_VOC_COEFF As Long = 9
Public Const EQUIPMENT_PANEL_LENGTH As Long = 10
Public Const EQUIPMENT_PANEL_WIDTH As Long = 11
Public Const EQUIPMENT_PANEL_COST As Long = 12
Public Const EQUIPMENT_INVERTER_MPPT_MIN As Long = 13
Public Const EQUIPMENT_INVERTER_MPPT_MAX As Long = 14
Public Const EQUIPMENT_INVERTER_VOC_MAX As Long = 15
Public Const EQUIPMENT_INVERTER_OPA_MAX As Long = 16
Public Const EQUIPMENT_INVERTER_ISC_MAX As Long = 17
Public Const EQUIPMENT_INVERTER_MPPT_COUNT As Long = 18
Public Const EQUIPMENT_INVERTER_MPPT_PORTS As Long = 19
Public Const EQUIPMENT_INVERTER_COST As Long = 20
Public Const EQUIPMENT_BATTERY_KW As Long = 21
Public Const EQUIPMENT_BATTERY_COST As Long = 22

'==============================================================
' Build Live Cells
'==============================================================
Public Const COL_LIVE_SUMMARY As Long = 6
Public Const LIVE_PANEL_QTY As Long = 4
Public Const LIVE_DC_ARRAY_KW As Long = 5
Public Const LIVE_STRING_VMP As Long = 6
Public Const LIVE_COLD_STRING_VOC As Long = 7
Public Const LIVE_MPPT_OP_A As Long = 8
Public Const LIVE_MPPT_SC_A As Long = 9
Public Const LIVE_KW_PER_MPPT As Long = 10
Public Const LIVE_BACKUP_BATTERY_KW As Long = 11
Public Const LIVE_ESTIMATED_HW_COST As Long = 12
Public Const LIVE_COST_PER_WATT As Long = 13
Public Const LIVE_ARRAY_WIDTH As Long = 16
Public Const LIVE_ARRAY_SLOPED_HEIGHT As Long = 17
Public Const LIVE_ARRAY_DEPTH As Long = 18
Public Const LIVE_ARRAY_REAR_EDGE As Long = 19

'==============================================================
' Workbook Named Ranges
'==============================================================

Public Const NAME_PANEL_LIST As String = "PanelList"
Public Const NAME_INVERTER_LIST As String = "InverterList"
Public Const NAME_BATTERY_LIST As String = "BatteryList"
Public Const NAME_ORIENTATION_LIST As String = "OrientationList"
Public Const NAME_MOUNT_TYPE_LIST As String = "MountTypeList"
Public Const NAME_MPPT_MODE_LIST As String = "MpptModeList"

Public Const NAME_SELECTED_PANEL_ID As String = "SelectedPanelID"
Public Const NAME_SELECTED_INVERTER_ID As String = "SelectedInverterID"
Public Const NAME_SELECTED_BATTERY_ID As String = "SelectedBatteryID"
Public Const NAME_PANELS_PER_STRING As String = "PanelsPerString"

'==============================================================
' Helpers
'==============================================================

Public Function BuildSheet() As Worksheet
    Set BuildSheet = ThisWorkbook.Worksheets(WS_BUILD)
End Function

Public Function DBListsSheet() As Worksheet
    Set DBListsSheet = ThisWorkbook.Worksheets(WS_DBLISTS)
End Function

Public Function BuildInputRange() As Range
    Set BuildInputRange = BuildSheet().Range( _
        BuildSheet().Cells(BUILD_INPUT_FIRST_ROW, COL_DESIGN_INPUTS), _
        BuildSheet().Cells(BUILD_INPUT_LAST_ROW, COL_DESIGN_INPUTS))
End Function

Public Function SelectedPanelIDCell() As Range
    Set SelectedPanelIDCell = DBListsSheet.Cells(ROW_SELECTED_PANEL, COL_SELECTED_ID_VALUE)
End Function

Public Function SelectedInverterIDCell() As Range
    Set SelectedInverterIDCell = DBListsSheet.Cells(ROW_SELECTED_INVERTER, COL_SELECTED_ID_VALUE)
End Function

Public Function SelectedBatteryIDCell() As Range
    Set SelectedBatteryIDCell = DBListsSheet.Cells(ROW_SELECTED_BATTERY, COL_SELECTED_ID_VALUE)
End Function

Public Function DesignPanelCell() As Range
    Set DesignPanelCell = _
        BuildSheet().Cells(BUILD_PANEL_CELL, COL_DESIGN_INPUTS)
End Function

Public Function DesignPanel() As String
    DesignPanel = CStr(DesignPanelCell().Value2)
End Function

Public Function DesignInverterCell() As Range
    Set DesignInverterCell = _
        BuildSheet().Cells(BUILD_INVERTER_CELL, COL_DESIGN_INPUTS)
End Function

Public Function DesignInverter() As String
    DesignInverter = CStr(DesignInverterCell().Value2)
End Function

Public Function DesignBatteryCell() As Range
    Set DesignBatteryCell = _
        BuildSheet().Cells(BUILD_BATTERY_CELL, COL_DESIGN_INPUTS)
End Function

Public Function DesignBattery() As String
    DesignBattery = CStr(DesignBatteryCell().Value2)
End Function

Public Function DesignBatteryQtyCell() As Range
    Set DesignBatteryQtyCell = _
        BuildSheet().Cells(BUILD_BATTERY_QTY, COL_DESIGN_INPUTS)
End Function

Public Function DesignBatteryQty() As Double
    DesignBatteryQty = CDbl(DesignBatteryQtyCell().Value2)
End Function

Public Function DesignBankTiltCell() As Range
    Set DesignBankTiltCell = _
        BuildSheet().Cells(BUILD_BANK_TILT, COL_DESIGN_INPUTS)
End Function

Public Function DesignBankTilt() As Double
    DesignBankTilt = CDbl(DesignBankTiltCell().Value2)
End Function

Public Function BatteryKwCell() As Range
    Set BatteryKwCell = _
        BuildSheet().Cells(EQUIPMENT_BATTERY_KW, COL_EQUIPMENT_VALUES)
End Function

Public Function BatteryKw() As Double
    BatteryKw = CDbl(BatteryKwCell().Value2)
End Function

Public Function DesignPanelsPerStringCell() As Range
    Set DesignPanelsPerStringCell = _
        BuildSheet().Cells(BUILD_PANELS_PER_STRING, COL_DESIGN_INPUTS)
End Function

Public Function DesignPanelsPerString() As Double
    DesignPanelsPerString = CDbl(DesignPanelsPerStringCell().Value2)
End Function

Public Function DesignTotalStringsCell() As Range
    Set DesignTotalStringsCell = _
        BuildSheet().Cells(BUILD_TOTAL_STRINGS, COL_DESIGN_INPUTS)
End Function

Public Function DesignTotalStrings() As Double
    DesignTotalStrings = CDbl(DesignTotalStringsCell().Value2)
End Function

Public Function DesignStringsPerMpptCell() As Range
    Set DesignStringsPerMpptCell = _
        BuildSheet().Cells(BUILD_STRINGS_PER_MPPT, COL_DESIGN_INPUTS)
End Function

Public Function DesignStringsPerMppt() As Double
    DesignStringsPerMppt = CDbl(DesignStringsPerMpptCell().Value2)
End Function

Public Function DesignMinTempCell() As Range
    Set DesignMinTempCell = _
        BuildSheet().Cells(BUILD_DESIGN_MIN_TEMP_C, COL_DESIGN_INPUTS)
End Function

Public Function DesignMinTemperature() As Double
    DesignMinTemperature = CDbl(DesignMinTempCell().Value2)
End Function

Public Function DesignMaxTempCell() As Range
    Set DesignMaxTempCell = _
        BuildSheet().Cells(BUILD_DESIGN_MAX_TEMP_C, COL_DESIGN_INPUTS)
End Function

Public Function DesignMaxTemperature() As Double
    DesignMaxTemperature = CDbl(DesignMaxTempCell().Value2)
End Function

Public Function DesignBankHeightCell() As Range
    Set DesignBankHeightCell = _
        BuildSheet().Cells(BUILD_BANK_HEIGHT, COL_DESIGN_INPUTS)
End Function

Public Function DesignBankHeight() As Double
    DesignBankHeight = CDbl(DesignBankHeightCell().Value2)
End Function

Public Function DesignBankWidthCell() As Range
    Set DesignBankWidthCell = _
        BuildSheet().Cells(BUILD_BANK_WIDTH, COL_DESIGN_INPUTS)
End Function

Public Function DesignBankWidth() As Double
    DesignBankWidth = CDbl(DesignBankWidthCell().Value2)
End Function

Public Function DesignBankCountCell() As Range
    Set DesignBankCountCell = _
        BuildSheet().Cells(BUILD_BANK_COUNT, COL_DESIGN_INPUTS)
End Function

Public Function DesignBankCount() As Double
    DesignBankCount = CDbl(DesignBankCountCell().Value2)
End Function

Public Function EquipmentPanelWattsCell() As Range
    Set EquipmentPanelWattsCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_WATTS, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelWatts() As Double
    EquipmentPanelWatts = CDbl(EquipmentPanelWattsCell().Value2)
End Function

Public Function EquipmentPanelVocCell() As Range
    Set EquipmentPanelVocCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_VOC, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelVoc() As Double
    EquipmentPanelVoc = CDbl(EquipmentPanelVocCell().Value2)
End Function

Public Function EquipmentPanelVmpCell() As Range
    Set EquipmentPanelVmpCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_VMP, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelVmp() As Double
    EquipmentPanelVmp = CDbl(EquipmentPanelVmpCell().Value2)
End Function

Public Function BuildStringsPerMpptCell() As Range
    Set BuildStringsPerMpptCell = _
        BuildSheet().Cells(BUILD_STRINGS_PER_MPPT, COL_DESIGN_INPUTS)
End Function

Public Function BuildStringsPerMppt() As Double
    BuildStringsPerMppt = CDbl(BuildStringsPerMpptCell().Value2)
End Function

Public Function EquipmentPanelIscCell() As Range
    Set EquipmentPanelIscCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_ISC, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelIsc() As Double
    EquipmentPanelIsc = CDbl(EquipmentPanelIscCell().Value2)
End Function

Public Function EquipmentPanelImpCell() As Range
    Set EquipmentPanelImpCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_IMP, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelImp() As Double
    EquipmentPanelImp = CDbl(EquipmentPanelImpCell().Value2)
End Function

Public Function EquipmentPanelCostCell() As Range
    Set EquipmentPanelCostCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_COST, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelCost() As Double
    EquipmentPanelCost = CDbl(EquipmentPanelCostCell().Value2)
End Function

Public Function EquipmentBatteryCostCell() As Range
    Set EquipmentBatteryCostCell = _
        BuildSheet().Cells(EQUIPMENT_BATTERY_COST, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentBatteryCost() As Double
    EquipmentBatteryCost = CDbl(EquipmentBatteryCostCell().Value2)
End Function

Public Function EquipmentInverterCostCell() As Range
    Set EquipmentInverterCostCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_COST, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInverterCost() As Double
    EquipmentInverterCost = CDbl(EquipmentInverterCostCell().Value2)
End Function

Public Function DesignBosCostCell() As Range
    Set DesignBosCostCell = _
        BuildSheet().Cells(BUILD_BOS_ESTIMATE, COL_DESIGN_INPUTS)
End Function

Public Function DesignBosCost() As Double
    DesignBosCost = CDbl(DesignBosCostCell().Value2)
End Function

Public Function LivePanelQtyCell() As Range
    Set LivePanelQtyCell = _
        BuildSheet().Cells(LIVE_PANEL_QTY, COL_LIVE_SUMMARY)
End Function

Public Function LivePanelQty() As Double
    LivePanelQty = CDbl(LivePanelQtyCell().Value2)
End Function

Public Function LiveDcArrayKwCell() As Range
    Set LiveDcArrayKwCell = _
        BuildSheet().Cells(LIVE_DC_ARRAY_KW, COL_LIVE_SUMMARY)
End Function

Public Function LiveDcArrayKw() As Double
    LiveDcArrayKw = CDbl(LiveDcArrayKwCell().Value2)
End Function

Public Function LiveNormalStringVmpCell() As Range
    Set LiveNormalStringVmpCell = _
        BuildSheet().Cells(LIVE_STRING_VMP, COL_LIVE_SUMMARY)
End Function

Public Function LiveNormalStringVmp() As Double
    LiveNormalStringVmp = CDbl(LiveNormalStringVmpCell().Value2)
End Function

Public Function LiveColdStringVocCell() As Range
    Set LiveColdStringVocCell = _
        BuildSheet().Cells(LIVE_COLD_STRING_VOC, COL_LIVE_SUMMARY)
End Function

Public Function LiveColdStringVoc() As Double
    LiveColdStringVoc = CDbl(LiveColdStringVocCell().Value2)
End Function

Public Function EquipmentPanelLengthCell() As Range
    Set EquipmentPanelLengthCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_LENGTH, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelLength() As Double
    EquipmentPanelLength = CDbl(EquipmentPanelLengthCell().Value2)
End Function

Public Function EquipmentPanelWidthCell() As Range
    Set EquipmentPanelWidthCell = _
        BuildSheet().Cells(EQUIPMENT_PANEL_WIDTH, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentPanelWidth() As Double
    EquipmentPanelWidth = CDbl(EquipmentPanelWidthCell().Value2)
End Function

Public Function EquipmentInvMpptCountCell() As Range
    Set EquipmentInvMpptCountCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_MPPT_COUNT, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvMpptCount() As Double
    EquipmentInvMpptCount = CDbl(EquipmentInvMpptCountCell().Value2)
End Function

Public Function EquipmentInvMpptMinCell() As Range
    Set EquipmentInvMpptMinCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_MPPT_MIN, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvMpptMin() As Double
    EquipmentInvMpptMin = CDbl(EquipmentInvMpptMinCell().Value2)
End Function

Public Function EquipmentInvMpptMaxCell() As Range
    Set EquipmentInvMpptMaxCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_MPPT_MAX, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvMpptMax() As Double
    EquipmentInvMpptMax = CDbl(EquipmentInvMpptMaxCell().Value2)
End Function

Public Function EquipmentInvVocMaxCell() As Range
    Set EquipmentInvVocMaxCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_VOC_MAX, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvVocMax() As Double
    EquipmentInvVocMax = CDbl(EquipmentInvVocMaxCell().Value2)
End Function

Public Function EquipmentInvOpAMaxCell() As Range
    Set EquipmentInvOpAMaxCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_OPA_MAX, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvOpAMax() As Double
    EquipmentInvOpAMax = CDbl(EquipmentInvOpAMaxCell().Value2)
End Function

Public Function EquipmentInvIscMaxCell() As Range
    Set EquipmentInvIscMaxCell = _
        BuildSheet().Cells(EQUIPMENT_INVERTER_ISC_MAX, COL_EQUIPMENT_VALUES)
End Function

Public Function EquipmentInvIscMax() As Double
    EquipmentInvIscMax = CDbl(EquipmentInvIscMaxCell().Value2)
End Function

Public Function DesignOrientationCell() As Range
    Set DesignOrientationCell = _
        BuildSheet().Cells(BUILD_ORIENTATION, COL_DESIGN_INPUTS)
End Function

Public Function DesignOrientation() As String
    DesignOrientation = CStr(DesignOrientationCell().Value2)
End Function

Public Function DesignMountTypeCell() As Range
    Set DesignMountTypeCell = _
        BuildSheet().Cells(BUILD_MOUNT_TYPE, COL_DESIGN_INPUTS)
End Function

Public Function DesignMountType() As String
    DesignMountType = CStr(DesignMountTypeCell().Value2)
End Function

Public Function DesignMpptModeCell() As Range
    Set DesignMpptModeCell = _
        BuildSheet().Cells(BUILD_MPPT_MODE, COL_DESIGN_INPUTS)
End Function

Public Function LiveArraySlopedHeightCell() As Range
    Set LiveArraySlopedHeightCell = _
        BuildSheet().Cells(LIVE_ARRAY_SLOPED_HEIGHT, COL_LIVE_SUMMARY)
End Function

Public Function LiveArraySlopedHeight() As Double
    LiveArraySlopedHeight = CDbl(LiveArraySlopedHeightCell().Value2)
End Function

Public Function LiveMpptOpACell() As Range
    Set LiveMpptOpACell = _
        BuildSheet().Cells(LIVE_MPPT_OP_A, COL_LIVE_SUMMARY)
End Function

Public Function LiveMpptOpA() As Double
    LiveMpptOpA = CDbl(LiveMpptOpACell().Value2)
End Function

Public Function LiveMpptScACell() As Range
    Set LiveMpptScACell = _
        BuildSheet().Cells(LIVE_MPPT_SC_A, COL_LIVE_SUMMARY)
End Function

Public Function LiveMpptScA() As Double
    LiveMpptScA = CDbl(LiveMpptScACell().Value2)
End Function

Public Function LiveMpptMinCell() As Range
    Set LiveMpptMinCell = _
        BuildSheet().Cells(LIVE_MPPT_, COL_LIVE_SUMMARY)
End Function

Public Function LiveMpptMin() As Double
    LiveMpptMin = CDbl(LiveMpptMinCell().Value2)
End Function


'************************
' Utilities
'************************

Public Sub SetOrReplaceName(ByVal nameText As String, ByVal refersToText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersToText
End Sub

Public Sub InitializeWorkbookLayoutNames()

    With DBListsSheet
        .Cells(ROW_SELECTED_PANEL, COL_SELECTED_ID_LABEL).value = "Selected Panel ID"
        .Cells(ROW_SELECTED_INVERTER, COL_SELECTED_ID_LABEL).value = "Selected Inverter ID"
        .Cells(ROW_SELECTED_BATTERY, COL_SELECTED_ID_LABEL).value = "Selected Battery ID"
    End With
        

    SetOrReplaceName NAME_SELECTED_PANEL_ID, _
        "=" & WS_DBLISTS & "!" & SelectedPanelIDCell.Address(True, True)

    SetOrReplaceName NAME_SELECTED_INVERTER_ID, _
        "=" & WS_DBLISTS & "!" & SelectedInverterIDCell.Address(True, True)

    SetOrReplaceName NAME_SELECTED_BATTERY_ID, _
        "=" & WS_DBLISTS & "!" & SelectedBatteryIDCell.Address(True, True)

    SetOrReplaceName NAME_PANELS_PER_STRING, _
        "=" & WS_BUILD & "!" & DesignPanelsPerStringCell.Address(True, True)

End Sub

Public Sub InitializeWorkbook()
    UnprotectBuildSheet
    InitializeEngine
    ProtectBuildSheet
End Sub

Public Sub UnprotectBuildSheet()
    BuildSheet().Unprotect
End Sub

Public Sub ProtectBuildSheet()

    If Not PROTECT_WORKBOOK Then Exit Sub

    With BuildSheet()

        .Unprotect

        .Cells.Locked = True

        BuildInputRange().Locked = False

        .Protect _
            DrawingObjects:=True, _
            Contents:=True, _
            Scenarios:=True

    End With

End Sub

Public Function LayoutNote() As String
'IF($B$10="Single Bank","Row spacing N/A - single structure","Check row spacing on Layout sheet")
    Dim mounttype As String
    Dim note As String
    
    mounttype = modWorkbookLayout.DesignMountType()
    
    If mounttype = "Single Bank" Then
        note = "Row spacing N/A - single structure"
    Else
        note = "Check row spacing on Layout sheet"
    End If
    
    LayoutNote = CStr(note)
    

End Function
