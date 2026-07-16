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
    Public Declare PtrSafe Function sdep_get_panels_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_source_notes Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#End If

Private Function PNum(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As Double
    Dim status As Long, value As Double
    EnsureDllInitialized
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
    EnsureDllInitialized
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
    EnsureDllInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_ids(StrPtr(buf), Len(buf))
    PanelIDs = ReadUtf16String(status, buf, "PanelIDs")
End Function

Public Function PanelModels() As String
    Dim status As Long, buf As String
    EnsureDllInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_panels_model_list(StrPtr(buf), Len(buf))
    PanelModels = ReadUtf16String(status, buf, "PanelModels")
End Function

Public Function PanelVmpByID(ByVal PanelID As String) As Double
    PanelVmpByID = PNum(PanelID, 1, "PanelVmpByID")
End Function

Public Function PanelVocByID(ByVal PanelID As String) As Double
    PanelVocByID = PNum(PanelID, 2, "PanelVocByID")
End Function

Public Function PanelImpByID(ByVal PanelID As String) As Double
    PanelImpByID = PNum(PanelID, 3, "PanelImpByID")
End Function

Public Function PanelIscByID(ByVal PanelID As String) As Double
    PanelIscByID = PNum(PanelID, 4, "PanelIscByID")
End Function

Public Function PanelWattsByID(ByVal PanelID As String) As Double
    PanelWattsByID = PNum(PanelID, 5, "PanelWattsByID")
End Function

Public Function PanelVocTempCoeffByID(ByVal PanelID As String) As Double
    PanelVocTempCoeffByID = PNum(PanelID, 6, "PanelVocTempCoeffByID")
End Function

Public Function PanelPmaxTempCoeffByID(ByVal PanelID As String) As Double
    PanelPmaxTempCoeffByID = PNum(PanelID, 7, "PanelPmaxTempCoeffByID")
End Function

Public Function PanelWidthMmByID(ByVal PanelID As String) As Double
    PanelWidthMmByID = PNum(PanelID, 8, "PanelWidthMmByID")
End Function

Public Function PanelWidthFtByID(ByVal PanelID As String) As Double
    PanelWidthFtByID = (PNum(PanelID, 8, "PanelWidthMmByID")) / 304.8
End Function

Public Function PanelLengthMmByID(ByVal PanelID As String) As Double
    PanelLengthMmByID = PNum(PanelID, 9, "PanelLengthMmByID")
End Function

Public Function PanelLengthFtByID(ByVal PanelID As String) As Double
    PanelLengthFtByID = (PNum(PanelID, 9, "PanelLengthMmByID")) / 304.8
End Function

Public Function PanelWeightLbByID(ByVal PanelID As String) As Double
    PanelWeightLbByID = PNum(PanelID, 10, "PanelWeightLbByID")
End Function

Public Function PanelSeriesFuseAByID(ByVal PanelID As String) As Double
    PanelSeriesFuseAByID = PNum(PanelID, 11, "PanelSeriesFuseAByID")
End Function

Public Function PanelMaxSystemVByID(ByVal PanelID As String) As Double
    PanelMaxSystemVByID = PNum(PanelID, 12, "PanelMaxSystemVByID")
End Function

Public Function PanelBallparkCostByID(ByVal PanelID As String) As Double
    PanelBallparkCostByID = PNum(PanelID, 13, "PanelBallparkCostByID")
End Function

Public Function PanelManufacturerByID(ByVal PanelID As String) As String
    PanelManufacturerByID = PStr(PanelID, 1, "PanelManufacturerByID")
End Function

Public Function PanelModelByID(ByVal PanelID As String) As String
    PanelModelByID = PStr(PanelID, 2, "PanelModelByID")
End Function

Public Function PanelSourceNotesByID(ByVal PanelID As String) As String
    PanelSourceNotesByID = PStr(PanelID, 3, "PanelSourceNotesByID")
End Function

Public Function PanelVmp() As Double
    PanelVmp = PanelVmpByID(SelectedPanelIDCell)
End Function

Public Function PanelVoc() As Double
    PanelVoc = PanelVocByID(SelectedPanelIDCell)
    End Function

Public Function PanelImp() As Double
    PanelImp = PanelImpByID(SelectedPanelIDCell)
End Function

Public Function PanelIsc() As Double
    PanelIsc = PanelIscByID(SelectedPanelIDCell)
End Function

Public Function panelwatts() As Double
    panelwatts = PanelWattsByID(SelectedPanelIDCell)
End Function

Public Function PanelVocTempCoeff() As Double
    PanelVocTempCoeff = PanelVocTempCoeffByID(SelectedPanelIDCell)
End Function

Public Function PanelPmaxTempCoeff() As Double
    PanelPmaxTempCoeff = PanelPmaxTempCoeffByID(SelectedPanelIDCell)
End Function

Public Function PanelWidthMm() As Double
    PanelWidthMm = PanelWidthMmByID(SelectedPanelIDCell)
End Function

Public Function PanelLengthMm() As Double
    PanelLengthMm = PanelLengthMmByID(SelectedPanelIDCell)
End Function

Public Function PanelWeightLb() As Double
    PanelWeightLb = PanelWeightLbByID(SelectedPanelIDCell)
End Function

Public Function PanelSeriesFuseA() As Double
    PanelSeriesFuseA = PanelSeriesFuseAByID(SelectedPanelIDCell)
End Function

Public Function PanelMaxSystemV() As Double
    PanelMaxSystemV = PanelMaxSystemVByID(SelectedPanelIDCell)
End Function

Public Function PanelBallparkCost() As Double
    PanelBallparkCost = PanelBallparkCostByID(SelectedPanelIDCell)
End Function

Public Function PanelManufacturer() As String
    PanelManufacturer = PanelManufacturerByID(SelectedPanelIDCell)
End Function

Public Function PanelModel() As String
    PanelModel = PanelModelByID(SelectedPanelIDCell)
End Function

Public Function PanelSourceNotes() As String
    PanelSourceNotes = PanelSourceNotesByID(SelectedPanelIDCell)
End Function

Public Function GetPanelWatts(ByVal PanelID As String) As Double
    GetPanelWatts = PNum(PanelID, 5, "GetPanelWatts")
End Function

Public Sub TestPanelIDs()
    MsgBox "Panel IDs:" & vbCrLf & PanelIDs(), vbInformation, "SDEP Panel ID Test"
End Sub

Public Sub TestPanelModels()
    MsgBox "Panel Models:" & vbCrLf & PanelModels(), vbInformation, "SDEP Panel Model Test"
End Sub

Public Sub TestPanelWatts()
    MsgBox "Panel Watts:" & vbCrLf & GetPanelWatts(SelectedPanelIDCell), vbInformation, "SDEP Panel Watts Test"
End Sub

Public Sub TestTotalPanels()
    MsgBox "Total Panels:" & vbCrLf & DesignTotalPanels(), vbInformation, "SDEP Total Panels Test"
End Sub


