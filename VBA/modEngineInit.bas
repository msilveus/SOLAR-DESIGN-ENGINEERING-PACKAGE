Attribute VB_Name = "modEngineInit"
Option Explicit

Public Const SDEP_SUCCESS As Long = 0
Public Const SDEP_BAD_ID As Long = 1
Public Const SDEP_BUFFER_TOO_SMALL As Long = 2
Public Const SDEP_INVALID_ARGUMENT As Long = 3
Public Const SDEP_DATABASE_NOT_LOADED As Long = 4
Public Const SDEP_DATABASE_CORRUPT As Long = 5
Public Const SDEP_WRONG_DATABASE_VERSION As Long = 6
Public Const SDEP_INTERNAL_ERROR As Long = 1000

Public Const SDEP_STRING_BUFFER_LEN As Long = 512

Private mEngineInitialized As Boolean
Private mDllInitialized As Boolean

#If VBA7 Then
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" _
        (ByVal lpPathName As String) As Long

    Public Declare PtrSafe Function sdep_init Lib "sdep_engine.dll" () As Long
#End If

Public Sub EnsureDllInitialized()

    Dim dllFolder As String
    Dim status As Long

    If mDllInitialized Then Exit Sub

    dllFolder = ThisWorkbook.Path

    If Len(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1000, "EnsureDllInitialized", _
            "Workbook must be saved before loading sdep_engine.dll."
    End If

    If Dir$(dllFolder & "\sdep_engine.dll") = "" Then
        Err.Raise vbObjectError + 1001, "EnsureDllInitialized", _
            "sdep_engine.dll not found in workbook folder: " & dllFolder
    End If

    If SetDllDirectory(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1002, "EnsureDllInitialized", _
            "Failed to set DLL search directory: " & dllFolder
    End If

    status = sdep_init()
    RaiseIfFailed status, "EnsureDllInitialized"

    mDllInitialized = True

End Sub

Public Sub InitializeEngine()

    Dim previousCalculation As XlCalculation
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    If mEngineInitialized Then Exit Sub

    On Error GoTo InitializationFailed

    With BuildInputRange()
        .Validation.Delete
    End With
    
    EnsureDllInitialized

    InitializeWorkbookLayoutNames
    
    BuildEquipmentDropdowns
    ApplyBuildOptionDropdowns
    ApplyEquipmentDropdowns

    SynchronizeSelectedEquipmentIDs

    mEngineInitialized = True

CleanExit:
    Exit Sub

InitializationFailed:
    mEngineInitialized = False

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub SynchronizeSelectedEquipmentIDs()

    Dim wsBuild As Worksheet

    Set wsBuild = BuildSheet()

    If Len(Trim$(CStr(DesignPanel()))) > 0 Then
        SaveSelectedID _
            selectedDisplay:=CStr(DesignPanel()), _
            displayName:=NAME_PANEL_LIST, _
            idColumn:=COL_PANEL_ID, _
            selectedIDName:=NAME_SELECTED_PANEL_ID
    End If

    If Len(Trim$(CStr(DesignInverter()))) > 0 Then
        SaveSelectedID _
            selectedDisplay:=CStr(DesignInverter()), _
            displayName:=NAME_INVERTER_LIST, _
            idColumn:=COL_INVERTER_ID, _
            selectedIDName:=NAME_SELECTED_INVERTER_ID
    End If

    If Len(Trim$(CStr(DesignBattery()))) > 0 Then
        SaveSelectedID _
            selectedDisplay:=CStr(DesignBattery()), _
            displayName:=NAME_BATTERY_LIST, _
            idColumn:=COL_BATTERY_ID, _
            selectedIDName:=NAME_SELECTED_BATTERY_ID
    End If

End Sub

Public Sub ResetEngineInitializationFlag()
    mEngineInitialized = False
End Sub

Public Sub EnsureEngineInitialized()
    If Not mEngineInitialized Then InitializeEngine
End Sub

Public Sub RaiseIfFailed(ByVal status As Long, ByVal sourceName As String)
    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, sourceName, _
            "SDEP engine call failed. Status code: " & CStr(status)
    End If
End Sub

Public Function TrimNullTerminatedString(ByVal s As String) As String
    Dim zeroPos As Long
    zeroPos = InStr(1, s, vbNullChar)
    If zeroPos > 0 Then
        TrimNullTerminatedString = Left$(s, zeroPos - 1)
    Else
        TrimNullTerminatedString = s
    End If
End Function

Public Function ReadUtf16String(ByVal status As Long, ByVal buf As String, ByVal sourceName As String) As String
    RaiseIfFailed status, sourceName
    ReadUtf16String = TrimNullTerminatedString(buf)
End Function

Public Sub InitializeWorkbookLayoutNames()

    With DBListsSheet()
        .Cells(ROW_SELECTED_PANEL, COL_SELECTED_ID_LABEL).value = "Selected Panel ID"
        .Cells(ROW_SELECTED_INVERTER, COL_SELECTED_ID_LABEL).value = "Selected Inverter ID"
        .Cells(ROW_SELECTED_BATTERY, COL_SELECTED_ID_LABEL).value = "Selected Battery ID"
    End With

    SetOrReplaceName NAME_SELECTED_PANEL_ID, _
        "=" & WS_DBLISTS & "!" & SelectedPanelIDCell().Address(True, True)

    SetOrReplaceName NAME_SELECTED_INVERTER_ID, _
        "=" & WS_DBLISTS & "!" & SelectedInverterIDCell().Address(True, True)

    SetOrReplaceName NAME_SELECTED_BATTERY_ID, _
        "=" & WS_DBLISTS & "!" & SelectedBatteryIDCell().Address(True, True)

End Sub
