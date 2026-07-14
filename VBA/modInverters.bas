Attribute VB_Name = "modInverters"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_get_inverters_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_mppt_min_v Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_mppt_max_v Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_pv_max_voc_v Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_max_operating_a_mppt Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_max_isc_a_mppt Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_mppt_count Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_strings_mppt Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_max_pv_kw Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_ac_output_kw Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_ballpark Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_inverters_manufacturer Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_model Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_notes Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#End If

Private Function INum(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    Select Case selector
        Case 1: status = sdep_get_inverters_mppt_min_v(id, value)
        Case 2: status = sdep_get_inverters_mppt_max_v(id, value)
        Case 3: status = sdep_get_inverters_pv_max_voc_v(id, value)
        Case 4: status = sdep_get_inverters_max_operating_a_mppt(id, value)
        Case 5: status = sdep_get_inverters_max_isc_a_mppt(id, value)
        Case 6: status = sdep_get_inverters_mppt_count(id, value)
        Case 7: status = sdep_get_inverters_strings_mppt(id, value)
        Case 8: status = sdep_get_inverters_max_pv_kw(id, value)
        Case 9: status = sdep_get_inverters_ac_output_kw(id, value)
        Case 10: status = sdep_get_inverters_ballpark(id, value)
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid inverter numeric selector."
    End Select
    RaiseIfFailed status, sourceName
    INum = value
End Function

Private Function IStr(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As String
    Dim status As Long, buf As String
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    Select Case selector
        Case 1: status = sdep_get_inverters_manufacturer(id, StrPtr(buf), Len(buf))
        Case 2: status = sdep_get_inverters_model(id, StrPtr(buf), Len(buf))
        Case 3: status = sdep_get_inverters_notes(id, StrPtr(buf), Len(buf))
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid inverter string selector."
    End Select
    IStr = ReadUtf16String(status, buf, sourceName)
End Function

Public Function InverterIDs() As String
    Dim status As Long, buf As String
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_inverters_ids(StrPtr(buf), Len(buf))
    InverterIDs = ReadUtf16String(status, buf, "InverterIDs")
End Function

Public Function InverterMpptMinV(ByVal InverterID As String) As Double: InverterMpptMinV = INum(InverterID, 1, "InverterMpptMinV"): End Function
Public Function InverterMpptMaxV(ByVal InverterID As String) As Double: InverterMpptMaxV = INum(InverterID, 2, "InverterMpptMaxV"): End Function
Public Function InverterMaxVoc(ByVal InverterID As String) As Double: InverterMaxVoc = INum(InverterID, 3, "InverterMaxVoc"): End Function
Public Function InverterMaxOperatingCurrent(ByVal InverterID As String) As Double: InverterMaxOperatingCurrent = INum(InverterID, 4, "InverterMaxOperatingCurrent"): End Function
Public Function InverterMaxIsc(ByVal InverterID As String) As Double: InverterMaxIsc = INum(InverterID, 5, "InverterMaxIsc"): End Function
Public Function InverterMpptCount(ByVal InverterID As String) As Double: InverterMpptCount = INum(InverterID, 6, "InverterMpptCount"): End Function
Public Function InverterStringsPerMppt(ByVal InverterID As String) As Double: InverterStringsPerMppt = INum(InverterID, 7, "InverterStringsPerMppt"): End Function
Public Function InverterMaxPvKw(ByVal InverterID As String) As Double: InverterMaxPvKw = INum(InverterID, 8, "InverterMaxPvKw"): End Function
Public Function InverterAcOutputKw(ByVal InverterID As String) As Double: InverterAcOutputKw = INum(InverterID, 9, "InverterAcOutputKw"): End Function
Public Function InverterBallparkCost(ByVal InverterID As String) As Double: InverterBallparkCost = INum(InverterID, 10, "InverterBallparkCost"): End Function
Public Function InverterManufacturer(ByVal InverterID As String) As String: InverterManufacturer = IStr(InverterID, 1, "InverterManufacturer"): End Function
Public Function InverterModel(ByVal InverterID As String) As String: InverterModel = IStr(InverterID, 2, "InverterModel"): End Function
Public Function InverterNotes(ByVal InverterID As String) As String: InverterNotes = IStr(InverterID, 3, "InverterNotes"): End Function

Public Sub TestInverterIDs()
    MsgBox "Inverter IDs:" & vbCrLf & InverterIDs(), vbInformation, "SDEP Inverter ID Test"
End Sub
