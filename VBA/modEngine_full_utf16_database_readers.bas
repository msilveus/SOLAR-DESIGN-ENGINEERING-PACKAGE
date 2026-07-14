Attribute VB_Name = "modEngine"
Option Explicit

' ============================================================
' SDEP Engine Bridge
' Excel/VBA front-end -> Rust sdep_engine.dll
'
' API convention:
'   - All DLL functions return SDEP_STATUS as Long
'   - Output numeric values are returned via ByRef parameters
'   - Output UTF-16 strings use caller-allocated VBA String buffers
'   - String buffers are passed as StrPtr(buffer)
'   - The DLL is expected to reside beside SDEP.xlsm
' ============================================================

Public Const SDEP_SUCCESS As Long = 0
Public Const SDEP_BAD_ID As Long = 1
Public Const SDEP_BUFFER_TOO_SMALL As Long = 2
Public Const SDEP_INVALID_ARGUMENT As Long = 3
Public Const SDEP_DATABASE_NOT_LOADED As Long = 4
Public Const SDEP_DATABASE_CORRUPT As Long = 5
Public Const SDEP_WRONG_DATABASE_VERSION As Long = 6
Public Const SDEP_INTERNAL_ERROR As Long = 1000

Private Const SDEP_STRING_BUFFER_LEN As Long = 512
Private mEngineInitialized As Boolean

#If VBA7 Then
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" _
        (ByVal lpPathName As String) As Long

    Public Declare PtrSafe Function sdep_init Lib "sdep_engine.dll" () As Long

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

    ' ---- Panel UTF-16 string readers ----
    Public Declare PtrSafe Function sdep_get_panels_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_source_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long

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

    ' ---- Inverter UTF-16 string readers ----
    Public Declare PtrSafe Function sdep_get_inverters_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long

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

    ' ---- Battery UTF-16 string readers ----
    Public Declare PtrSafe Function sdep_get_batteries_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_chemistry Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_closed_loop_with_sol_ark Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_comm_interface Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long

    ' ---- Existing simple calculation functions if still exported ----
    Public Declare PtrSafe Function sdep_string_vmp Lib "sdep_engine.dll" _
        (ByVal panel_vmp As Double, ByVal panels_per_string As Long) As Double
    Public Declare PtrSafe Function sdep_string_voc Lib "sdep_engine.dll" _
        (ByVal panel_voc As Double, ByVal panels_per_string As Long) As Double
#End If

Public Sub InitializeEngine()
    Dim dllFolder As String
    Dim status As Long

    If mEngineInitialized Then Exit Sub

    dllFolder = ThisWorkbook.Path

    If Len(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1000, "InitializeEngine", _
            "Workbook must be saved before loading sdep_engine.dll."
    End If

    If Dir(dllFolder & "\sdep_engine.dll") = "" Then
        Err.Raise vbObjectError + 1001, "InitializeEngine", _
            "sdep_engine.dll not found in workbook folder: " & dllFolder
    End If

    If SetDllDirectory(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1002, "InitializeEngine", _
            "Failed to set DLL search directory: " & dllFolder
    End If

    status = sdep_init()
    RaiseIfFailed status, "InitializeEngine"

    mEngineInitialized = True
End Sub

Public Sub ResetEngineInitializationFlag()
    mEngineInitialized = False
End Sub

Private Sub EnsureEngineInitialized()
    If Not mEngineInitialized Then InitializeEngine
End Sub

Private Sub RaiseIfFailed(ByVal status As Long, ByVal sourceName As String)
    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, sourceName, _
            "SDEP engine call failed. Status code: " & CStr(status)
    End If
End Sub

Private Function TrimNullTerminatedString(ByVal s As String) As String
    Dim zeroPos As Long
    zeroPos = InStr(1, s, vbNullChar)
    If zeroPos > 0 Then
        TrimNullTerminatedString = Left$(s, zeroPos - 1)
    Else
        TrimNullTerminatedString = s
    End If
End Function

Private Function ReadUtf16String(ByVal status As Long, ByVal buf As String, ByVal sourceName As String) As String
    RaiseIfFailed status, sourceName
    ReadUtf16String = TrimNullTerminatedString(buf)
End Function

' ============================================================
' Panel wrappers
' ============================================================

Public Function PanelVmp(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_vmp(PanelID, value)
    RaiseIfFailed status, "PanelVmp"
    PanelVmp = value
End Function

Public Function PanelVoc(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_voc(PanelID, value)
    RaiseIfFailed status, "PanelVoc"
    PanelVoc = value
End Function

Public Function PanelImp(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_imp(PanelID, value)
    RaiseIfFailed status, "PanelImp"
    PanelImp = value
End Function

Public Function PanelIsc(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_isc(PanelID, value)
    RaiseIfFailed status, "PanelIsc"
    PanelIsc = value
End Function

Public Function PanelWatts(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_stc_w(PanelID, value)
    RaiseIfFailed status, "PanelWatts"
    PanelWatts = value
End Function

Public Function PanelVocTempCoeff(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_voc_temp_c(PanelID, value)
    RaiseIfFailed status, "PanelVocTempCoeff"
    PanelVocTempCoeff = value
End Function

Public Function PanelPmaxTempCoeff(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_pmax_temp_c(PanelID, value)
    RaiseIfFailed status, "PanelPmaxTempCoeff"
    PanelPmaxTempCoeff = value
End Function

Public Function PanelWidthMm(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_width_mm(PanelID, value)
    RaiseIfFailed status, "PanelWidthMm"
    PanelWidthMm = value
End Function

Public Function PanelLengthMm(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_length_mm(PanelID, value)
    RaiseIfFailed status, "PanelLengthMm"
    PanelLengthMm = value
End Function

Public Function PanelWeightLb(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_weight_lb(PanelID, value)
    RaiseIfFailed status, "PanelWeightLb"
    PanelWeightLb = value
End Function

Public Function PanelSeriesFuseA(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_series_fuse_a(PanelID, value)
    RaiseIfFailed status, "PanelSeriesFuseA"
    PanelSeriesFuseA = value
End Function

Public Function PanelMaxSystemV(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_max_system_v(PanelID, value)
    RaiseIfFailed status, "PanelMaxSystemV"
    PanelMaxSystemV = value
End Function

Public Function PanelBallparkCost(ByVal PanelID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_panels_ballpark_panel(PanelID, value)
    RaiseIfFailed status, "PanelBallparkCost"
    PanelBallparkCost = value
End Function

Public Function PanelManufacturer(ByVal PanelID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_manufacturer(PanelID, StrPtr(buf), Len(buf))
    PanelManufacturer = ReadUtf16String(status, buf, "PanelManufacturer")
End Function

Public Function PanelModel(ByVal PanelID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_model(PanelID, StrPtr(buf), Len(buf))
    PanelModel = ReadUtf16String(status, buf, "PanelModel")
End Function

Public Function PanelSourceNotes(ByVal PanelID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_source_notes(PanelID, StrPtr(buf), Len(buf))
    PanelSourceNotes = ReadUtf16String(status, buf, "PanelSourceNotes")
End Function

' ============================================================
' Inverter wrappers
' ============================================================

Public Function InverterMpptMinV(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_mppt_min_v(InverterID, value)
    RaiseIfFailed status, "InverterMpptMinV"
    InverterMpptMinV = value
End Function

Public Function InverterMpptMaxV(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_mppt_max_v(InverterID, value)
    RaiseIfFailed status, "InverterMpptMaxV"
    InverterMpptMaxV = value
End Function

Public Function InverterMaxVoc(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_pv_max_voc_v(InverterID, value)
    RaiseIfFailed status, "InverterMaxVoc"
    InverterMaxVoc = value
End Function

Public Function InverterMaxOperatingCurrent(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_max_operating_a_mppt(InverterID, value)
    RaiseIfFailed status, "InverterMaxOperatingCurrent"
    InverterMaxOperatingCurrent = value
End Function

Public Function InverterMaxIsc(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_max_isc_a_mppt(InverterID, value)
    RaiseIfFailed status, "InverterMaxIsc"
    InverterMaxIsc = value
End Function

Public Function InverterMpptCount(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_mppt_count(InverterID, value)
    RaiseIfFailed status, "InverterMpptCount"
    InverterMpptCount = value
End Function

Public Function InverterStringsPerMppt(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_strings_mppt(InverterID, value)
    RaiseIfFailed status, "InverterStringsPerMppt"
    InverterStringsPerMppt = value
End Function

Public Function InverterMaxPvKw(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_max_pv_kw(InverterID, value)
    RaiseIfFailed status, "InverterMaxPvKw"
    InverterMaxPvKw = value
End Function

Public Function InverterAcOutputKw(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_ac_output_kw(InverterID, value)
    RaiseIfFailed status, "InverterAcOutputKw"
    InverterAcOutputKw = value
End Function

Public Function InverterBallparkCost(ByVal InverterID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_inverters_ballpark(InverterID, value)
    RaiseIfFailed status, "InverterBallparkCost"
    InverterBallparkCost = value
End Function

Public Function InverterManufacturer(ByVal InverterID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_inverters_manufacturer(InverterID, StrPtr(buf), Len(buf))
    InverterManufacturer = ReadUtf16String(status, buf, "InverterManufacturer")
End Function

Public Function InverterModel(ByVal InverterID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_inverters_model(InverterID, StrPtr(buf), Len(buf))
    InverterModel = ReadUtf16String(status, buf, "InverterModel")
End Function

Public Function InverterNotes(ByVal InverterID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_inverters_notes(InverterID, StrPtr(buf), Len(buf))
    InverterNotes = ReadUtf16String(status, buf, "InverterNotes")
End Function

' ============================================================
' Battery wrappers
' ============================================================

Public Function BatteryUsableKwh(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_usable_kwh(BatteryID, value)
    RaiseIfFailed status, "BatteryUsableKwh"
    BatteryUsableKwh = value
End Function

Public Function BatteryNominalV(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_nominal_v(BatteryID, value)
    RaiseIfFailed status, "BatteryNominalV"
    BatteryNominalV = value
End Function

Public Function BatteryCapacityAh(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_capacity_ah(BatteryID, value)
    RaiseIfFailed status, "BatteryCapacityAh"
    BatteryCapacityAh = value
End Function

Public Function BatteryContChargeA(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_cont_charge_a(BatteryID, value)
    RaiseIfFailed status, "BatteryContChargeA"
    BatteryContChargeA = value
End Function

Public Function BatteryContDischargeA(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_cont_discharge_a(BatteryID, value)
    RaiseIfFailed status, "BatteryContDischargeA"
    BatteryContDischargeA = value
End Function

Public Function BatteryPeakA(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_peak_a(BatteryID, value)
    RaiseIfFailed status, "BatteryPeakA"
    BatteryPeakA = value
End Function

Public Function BatteryWeightLb(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_weight_lb(BatteryID, value)
    RaiseIfFailed status, "BatteryWeightLb"
    BatteryWeightLb = value
End Function

Public Function BatteryBallparkCost(ByVal BatteryID As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    status = sdep_get_batteries_ballpark_battery(BatteryID, value)
    RaiseIfFailed status, "BatteryBallparkCost"
    BatteryBallparkCost = value
End Function

Public Function BatteryManufacturer(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_manufacturer(BatteryID, StrPtr(buf), Len(buf))
    BatteryManufacturer = ReadUtf16String(status, buf, "BatteryManufacturer")
End Function

Public Function BatteryModel(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_model(BatteryID, StrPtr(buf), Len(buf))
    BatteryModel = ReadUtf16String(status, buf, "BatteryModel")
End Function

Public Function BatteryChemistry(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_chemistry(BatteryID, StrPtr(buf), Len(buf))
    BatteryChemistry = ReadUtf16String(status, buf, "BatteryChemistry")
End Function

Public Function BatteryClosedLoopWithSolArk(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_closed_loop_with_sol_ark(BatteryID, StrPtr(buf), Len(buf))
    BatteryClosedLoopWithSolArk = ReadUtf16String(status, buf, "BatteryClosedLoopWithSolArk")
End Function

Public Function BatteryCommInterface(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_comm_interface(BatteryID, StrPtr(buf), Len(buf))
    BatteryCommInterface = ReadUtf16String(status, buf, "BatteryCommInterface")
End Function

Public Function BatteryNotes(ByVal BatteryID As String) As String
    Dim buf As String, status As Long
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_notes(BatteryID, StrPtr(buf), Len(buf))
    BatteryNotes = ReadUtf16String(status, buf, "BatteryNotes")
End Function

Public Function EngineStringVmp(ByVal PanelVmpValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureEngineInitialized
    EngineStringVmp = sdep_string_vmp(PanelVmpValue, PanelsPerString)
End Function

Public Function EngineStringVoc(ByVal PanelVocValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureEngineInitialized
    EngineStringVoc = sdep_string_voc(PanelVocValue, PanelsPerString)
End Function

Public Sub TestEngineDatabaseReaders()
    Dim msg As String

    InitializeEngine

    msg = "Panel P000001:" & vbCrLf & _
          "  Manufacturer: " & PanelManufacturer("P000001") & vbCrLf & _
          "  Model: " & PanelModel("P000001") & vbCrLf & _
          "  Vmp: " & PanelVmp("P000001") & vbCrLf & _
          "  Voc: " & PanelVoc("P000001") & vbCrLf & vbCrLf & _
          "Inverter I000001:" & vbCrLf & _
          "  Manufacturer: " & InverterManufacturer("I000001") & vbCrLf & _
          "  Model: " & InverterModel("I000001") & vbCrLf & _
          "  MPPT Range: " & InverterMpptMinV("I000001") & " - " & InverterMpptMaxV("I000001") & " V" & vbCrLf & vbCrLf & _
          "Battery B000001:" & vbCrLf & _
          "  Manufacturer: " & BatteryManufacturer("B000001") & vbCrLf & _
          "  Model: " & BatteryModel("B000001") & vbCrLf & _
          "  Chemistry: " & BatteryChemistry("B000001") & vbCrLf & _
          "  Usable kWh: " & BatteryUsableKwh("B000001")

    MsgBox msg, vbInformation, "SDEP Engine Database Reader Test"
End Sub
