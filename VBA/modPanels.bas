Attribute VB_Name = "modPanels"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_get_panels_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_vmp Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_voc Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_imp Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_isc Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_stc_w Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_voc_temp_c Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_pmax_temp_c Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_width_mm Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_length_mm Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_weight_lb Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_series_fuse_a Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_max_system_v Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_ballpark_panel Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_panels_manufacturer Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_model Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_source_notes Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#End If

Private Function PNum(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As Double
    Dim status As Long, value As Double
    EnsureEngineInitialized
    Select Case selector
        Case 1: status = sdep_get_panels_vmp(id, value)
        Case 2: status = sdep_get_panels_voc(id, value)
        Case 3: status = sdep_get_panels_imp(id, value)
        Case 4: status = sdep_get_panels_isc(id, value)
        Case 5: status = sdep_get_panels_stc_w(id, value)
        Case 6: status = sdep_get_panels_voc_temp_c(id, value)
        Case 7: status = sdep_get_panels_pmax_temp_c(id, value)
        Case 8: status = sdep_get_panels_width_mm(id, value)
        Case 9: status = sdep_get_panels_length_mm(id, value)
        Case 10: status = sdep_get_panels_weight_lb(id, value)
        Case 11: status = sdep_get_panels_series_fuse_a(id, value)
        Case 12: status = sdep_get_panels_max_system_v(id, value)
        Case 13: status = sdep_get_panels_ballpark_panel(id, value)
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid panel numeric selector."
    End Select
    RaiseIfFailed status, sourceName
    PNum = value
End Function

Private Function PStr(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As String
    Dim status As Long, buf As String
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    Select Case selector
        Case 1: status = sdep_get_panels_manufacturer(id, StrPtr(buf), Len(buf))
        Case 2: status = sdep_get_panels_model(id, StrPtr(buf), Len(buf))
        Case 3: status = sdep_get_panels_source_notes(id, StrPtr(buf), Len(buf))
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid panel string selector."
    End Select
    PStr = ReadUtf16String(status, buf, sourceName)
End Function

Public Function PanelIDs() As String
    Dim status As Long, buf As String
    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_ids(StrPtr(buf), Len(buf))
    PanelIDs = ReadUtf16String(status, buf, "PanelIDs")
End Function

Public Function PanelVmp(ByVal PanelID As String) As Double: PanelVmp = PNum(PanelID, 1, "PanelVmp"): End Function
Public Function PanelVoc(ByVal PanelID As String) As Double: PanelVoc = PNum(PanelID, 2, "PanelVoc"): End Function
Public Function PanelImp(ByVal PanelID As String) As Double: PanelImp = PNum(PanelID, 3, "PanelImp"): End Function
Public Function PanelIsc(ByVal PanelID As String) As Double: PanelIsc = PNum(PanelID, 4, "PanelIsc"): End Function
Public Function PanelWatts(ByVal PanelID As String) As Double: PanelWatts = PNum(PanelID, 5, "PanelWatts"): End Function
Public Function PanelVocTempCoeff(ByVal PanelID As String) As Double: PanelVocTempCoeff = PNum(PanelID, 6, "PanelVocTempCoeff"): End Function
Public Function PanelPmaxTempCoeff(ByVal PanelID As String) As Double: PanelPmaxTempCoeff = PNum(PanelID, 7, "PanelPmaxTempCoeff"): End Function
Public Function PanelWidthMm(ByVal PanelID As String) As Double: PanelWidthMm = PNum(PanelID, 8, "PanelWidthMm"): End Function
Public Function PanelLengthMm(ByVal PanelID As String) As Double: PanelLengthMm = PNum(PanelID, 9, "PanelLengthMm"): End Function
Public Function PanelWeightLb(ByVal PanelID As String) As Double: PanelWeightLb = PNum(PanelID, 10, "PanelWeightLb"): End Function
Public Function PanelSeriesFuseA(ByVal PanelID As String) As Double: PanelSeriesFuseA = PNum(PanelID, 11, "PanelSeriesFuseA"): End Function
Public Function PanelMaxSystemV(ByVal PanelID As String) As Double: PanelMaxSystemV = PNum(PanelID, 12, "PanelMaxSystemV"): End Function
Public Function PanelBallparkCost(ByVal PanelID As String) As Double: PanelBallparkCost = PNum(PanelID, 13, "PanelBallparkCost"): End Function
Public Function PanelManufacturer(ByVal PanelID As String) As String: PanelManufacturer = PStr(PanelID, 1, "PanelManufacturer"): End Function
Public Function PanelModel(ByVal PanelID As String) As String: PanelModel = PStr(PanelID, 2, "PanelModel"): End Function
Public Function PanelSourceNotes(ByVal PanelID As String) As String: PanelSourceNotes = PStr(PanelID, 3, "PanelSourceNotes"): End Function

Public Sub TestPanelIDs()
    MsgBox "Panel IDs:" & vbCrLf & PanelIDs(), vbInformation, "SDEP Panel ID Test"
End Sub
