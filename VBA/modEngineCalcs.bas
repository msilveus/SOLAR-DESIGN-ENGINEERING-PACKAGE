Attribute VB_Name = "modEngineCalcs"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_string_vmp Lib "sdep_engine.dll" (ByVal panel_vmp As Double, ByVal panels_per_string As Long) As Double
    Public Declare PtrSafe Function sdep_string_voc Lib "sdep_engine.dll" (ByVal panel_voc As Double, ByVal panels_per_string As Long) As Double
#End If

Public Function EngineStringVmp(ByVal PanelVmpValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureEngineInitialized
    EngineStringVmp = sdep_string_vmp(PanelVmpValue, PanelsPerString)
End Function

Public Function EngineStringVoc(ByVal PanelVocValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureEngineInitialized
    EngineStringVoc = sdep_string_voc(PanelVocValue, PanelsPerString)
End Function
