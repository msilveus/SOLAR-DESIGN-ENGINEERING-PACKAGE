Attribute VB_Name = "modEngine"
Option Explicit

' ============================================================
' SDEP Engine Bridge
' Excel/VBA front-end -> Rust sdep_engine.dll
'
' API convention:
'   - All DLL functions return SDEP_STATUS as Long
'   - Output numeric values are returned via ByRef parameters
'   - Output strings use caller-allocated buffers
'   - The DLL is expected to reside beside SDEP.xlsm
' ============================================================

' ----------------------------
' Status codes
' ----------------------------
Public Const SDEP_SUCCESS As Long = 0
Public Const SDEP_BAD_ID As Long = 1
Public Const SDEP_BUFFER_TOO_SMALL As Long = 2
Public Const SDEP_INVALID_ARGUMENT As Long = 3
Public Const SDEP_DATABASE_NOT_LOADED As Long = 4
Public Const SDEP_DATABASE_CORRUPT As Long = 5
Public Const SDEP_WRONG_DATABASE_VERSION As Long = 6
Public Const SDEP_INTERNAL_ERROR As Long = 1000

Private Const SDEP_STRING_BUFFER_LEN As Long = 512

#If VBA7 Then

    Public Declare PtrSafe Function sdep_init Lib "sdep_engine.dll" () As Long
    
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" _
        (ByVal lpPathName As String) As Long

    ' ---- Panel numeric readers ----
    Public Declare PtrSafe Function sdep_get_panels_vmp Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_voc Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_imp Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_isc Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_stc_w Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_voc_temp_c Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_pmax_temp_c Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_width_mm Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_length_mm Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_weight_lb Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_series_fuse_a Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_max_system_v Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_panels_ballpark_panel Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    ' ---- Panel string readers ----
    Public Declare PtrSafe Function sdep_get_panels_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    
    Public Declare PtrSafe Function sdep_get_panels_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef buffer As Byte, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_panels_source_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    ' ---- Inverter numeric readers ----
    Public Declare PtrSafe Function sdep_get_inverters_mppt_min_v Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_mppt_max_v Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_pv_max_voc_v Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_max_operating_a_mppt Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_max_isc_a_mppt Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_mppt_count Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_strings_mppt Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_max_pv_kw Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_ac_output_kw Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_inverters_ballpark Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    ' ---- Inverter string readers ----
    Public Declare PtrSafe Function sdep_get_inverters_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_inverters_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_inverters_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    ' ---- Battery numeric readers ----
    Public Declare PtrSafe Function sdep_get_batteries_usable_kwh Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_nominal_v Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_capacity_ah Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_cont_charge_a Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_cont_discharge_a Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_peak_a Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_weight_lb Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    Public Declare PtrSafe Function sdep_get_batteries_ballpark_battery Lib "sdep_engine.dll" _
        (ByVal id As String, ByRef out_value As Double) As Long

    ' ---- Battery string readers ----
    Public Declare PtrSafe Function sdep_get_batteries_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_chemistry Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_closed_loop_with_sol_ark Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_comm_interface Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As String, ByVal buffer_len As LongPtr) As Long

    ' ---- Existing calculation functions if still exported ----
    Public Declare PtrSafe Function sdep_string_vmp Lib "sdep_engine.dll" _
        (ByVal panel_vmp As Double, ByVal panels_per_string As Long) As Double

    Public Declare PtrSafe Function sdep_string_voc Lib "sdep_engine.dll" _
        (ByVal panel_voc As Double, ByVal panels_per_string As Long) As Double

#End If

' ============================================================
' Engine initialization
' ============================================================

Private mEngineInitialized As Boolean

Public Sub InitializeEngine()

    Dim dllFolder As String
    Dim dataJsonFolder As String
    Dim status As Long

    If mEngineInitialized Then Exit Sub

    dllFolder = ThisWorkbook.Path
    dataJsonFolder = ThisWorkbook.Path & "\DataJson"

    If Len(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1000, "InitializeEngine", _
            "Workbook must be saved before loading sdep_engine.dll."
    End If

    If Dir(dllFolder & "\sdep_engine.dll") = "" Then
        Err.Raise vbObjectError + 1001, "InitializeEngine", _
            "sdep_engine.dll not found in workbook folder: " & dllFolder
    End If

    If Dir(dataJsonFolder, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1002, "InitializeEngine", _
            "DataJson folder not found: " & dataJsonFolder
    End If

    If SetDllDirectory(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1003, "InitializeEngine", _
            "Failed to set DLL search directory: " & dllFolder
    End If

    status = sdep_init()

    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, "InitializeEngine", _
            "SDEP engine initialization failed. Status code: " & CStr(status)
    End If

    mEngineInitialized = True

End Sub

' ============================================================
' Status handling
' ============================================================

Private Sub RaiseIfFailed(ByVal status As Long, ByVal sourceName As String)
    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, sourceName, _
            "SDEP engine call failed. Status code: " & CStr(status)
    End If
End Sub

Private Function ReadStringFromEngine(ByVal status As Long, ByVal buf As String, ByVal sourceName As String) As String
    Dim zeroPos As Long

    RaiseIfFailed status, sourceName

    zeroPos = InStr(1, buf, vbNullChar)
    If zeroPos > 0 Then
        ReadStringFromEngine = Left$(buf, zeroPos - 1)
    Else
        ReadStringFromEngine = buf
    End If
End Function

Private Function ReadEngineString(ByVal status As Long, ByRef bytes() As Byte) As String
    Dim i As Long

    RaiseIfFailed status, "ReadEngineString"

    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) = 0 Then Exit For
    Next i

    ReadEngineString = StrConv(LeftB$(bytes, i), vbUnicode)
End Function

Private Function BytesToString(ByRef bytes() As Byte) As String
    Dim i As Long
    Dim n As Long

    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) = 0 Then Exit For
    Next i

    n = i - LBound(bytes)
    If n <= 0 Then
        BytesToString = ""
    Else
        ReDim Preserve bytes(0 To n - 1)
        BytesToString = StrConv(bytes, vbUnicode)
    End If
End Function

Private Function TrimNullTerminatedString(ByVal s As String) As String
    Dim zeroPos As Long
    zeroPos = InStr(1, s, vbNullChar)

    If zeroPos > 0 Then
        TrimNullTerminatedString = Left$(s, zeroPos - 1)
    Else
        TrimNullTerminatedString = s
    End If
End Function

' ============================================================
' Panel wrappers
' ============================================================

Public Function PanelVmp(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_vmp(PanelID, value)
    RaiseIfFailed status, "PanelVmp"
    PanelVmp = value
End Function

Public Function PanelVoc(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_voc(PanelID, value)
    RaiseIfFailed status, "PanelVoc"
    PanelVoc = value
End Function

Public Function PanelImp(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_imp(PanelID, value)
    RaiseIfFailed status, "PanelImp"
    PanelImp = value
End Function

Public Function PanelIsc(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_isc(PanelID, value)
    RaiseIfFailed status, "PanelIsc"
    PanelIsc = value
End Function

Public Function PanelWatts(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_stc_w(PanelID, value)
    RaiseIfFailed status, "PanelWatts"
    PanelWatts = value
End Function

Public Function PanelVocTempCoeff(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_panels_voc_temp_c(PanelID, value)
    RaiseIfFailed status, "PanelVocTempCoeff"
    PanelVocTempCoeff = value
End Function

Public Function PanelModel(ByVal PanelID As String) As String
    Dim buf As String, status As Long
    InitializeEngine
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_model(PanelID, buf, Len(buf))

    RaiseIfFailed status, "PanelModel"
    PanelModel = ReadStringFromEngine(status, buf, "PanelManufacturer")
End Function

Public Function PanelManufacturer(ByVal PanelID As String) As String
    Dim buf As String
    Dim status As Long

    InitializeEngine

    buf = String$(512, vbNullChar)

    status = sdep_get_panels_manufacturer(PanelID, StrPtr(buf), Len(buf))
    RaiseIfFailed status, "PanelManufacturer"

    PanelManufacturer = TrimNullTerminatedString(buf)
End Function
' ============================================================
' Inverter wrappers
' ============================================================

Public Function InverterMpptMinV(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_inverters_mppt_min_v(InverterID, value)
    RaiseIfFailed status, "InverterMpptMinV"
    InverterMpptMinV = value
End Function

Public Function InverterMpptMaxV(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_inverters_mppt_max_v(InverterID, value)
    RaiseIfFailed status, "InverterMpptMaxV"
    InverterMpptMaxV = value
End Function

Public Function InverterMaxVoc(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_inverters_pv_max_voc_v(InverterID, value)
    RaiseIfFailed status, "InverterMaxVoc"
    InverterMaxVoc = value
End Function

Public Function InverterMaxOperatingCurrent(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_inverters_max_operating_a_mppt(InverterID, value)
    RaiseIfFailed status, "InverterMaxOperatingCurrent"
    InverterMaxOperatingCurrent = value
End Function

Public Function InverterMaxIsc(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_inverters_max_isc_a_mppt(InverterID, value)
    RaiseIfFailed status, "InverterMaxIsc"
    InverterMaxIsc = value
End Function

Public Function InverterModel(ByVal InverterID As String) As String
    Dim buf As String, status As Long
    InitializeEngine
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_inverters_model(InverterID, buf, Len(buf))
    InverterModel = ReadStringFromEngine(status, buf, "InverterModel")
End Function

' ============================================================
' Battery wrappers
' ============================================================

Public Function BatteryUsableKwh(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_batteries_usable_kwh(BatteryID, value)
    RaiseIfFailed status, "BatteryUsableKwh"
    BatteryUsableKwh = value
End Function

Public Function BatteryNominalV(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    InitializeEngine
    status = sdep_get_batteries_nominal_v(BatteryID, value)
    RaiseIfFailed status, "BatteryNominalV"
    BatteryNominalV = value
End Function

Public Function BatteryModel(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    InitializeEngine
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_model(BatteryID, buf, Len(buf))
    BatteryModel = ReadStringFromEngine(status, buf, "BatteryModel")
End Function

' ============================================================
' Simple tests
' ============================================================

Public Sub TestEngineDatabaseReaders()
    Dim msg As String

    InitializeEngine

    msg = "Panel P000001:" & vbCrLf & _
          "  Manufacturer: " & PanelManufacturer("P000001") & vbCrLf & _
          "  Model: " & PanelModel("P000001") & vbCrLf & _
          "  Vmp: " & PanelVmp("P000001") & vbCrLf & _
          "  Voc: " & PanelVoc("P000001") & vbCrLf & vbCrLf & _
          "Inverter I000001:" & vbCrLf & _
          "  Model: " & InverterModel("I000001") & vbCrLf & _
          "  MPPT Range: " & InverterMpptMinV("I000001") & " - " & InverterMpptMaxV("I000001") & " V" & vbCrLf & vbCrLf & _
          "Battery B000001:" & vbCrLf & _
          "  Model: " & BatteryModel("B000001") & vbCrLf & _
          "  Usable kWh: " & BatteryUsableKwh("B000001")

    MsgBox msg, vbInformation, "SDEP Engine Database Reader Test"
End Sub

