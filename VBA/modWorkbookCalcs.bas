Attribute VB_Name = "modWorkbookCalcs"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_string_vmp Lib "sdep_engine.dll" (ByVal panel_vmp As Double, ByVal panels_per_string As Long) As Double
    Public Declare PtrSafe Function sdep_string_voc Lib "sdep_engine.dll" (ByVal panel_voc As Double, ByVal panels_per_string As Long) As Double
#End If

Public Function EngineStringVmp(ByVal PanelVmpValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureDllInitialized
    EngineStringVmp = sdep_string_vmp(PanelVmpValue, PanelsPerString)
End Function

Public Function EngineStringVoc(ByVal PanelVocValue As Double, ByVal PanelsPerString As Long) As Double
    EnsureDllInitialized
    EngineStringVoc = sdep_string_voc(PanelVocValue, PanelsPerString)
End Function

Public Function ColdStringVoc() As Double

    ColdStringVoc = _
        PanelVoc() * _
        PanelsPerString() * _
        (1# + Abs(PanelVocTempCoeff()) / 100# * (25# - DesignTemperature()))

End Function

Public Function DesignTotalPanels() As Long
    Dim height As Long
    Dim width As Long
    Dim count As Long
    
    height = DesignBankHeight()
    width = DesignBankWidth()
    count = DesignBankCount()
    DesignTotalPanels = CLng(height * width * count)
End Function

Public Function LiveArrayKw() As Double
    Dim panelwatts As Double
    Dim totalpanels As Double
    Dim totalarraywatts As Double
    
    panelwatts = CDbl(EquipmentPanelWatts())
    totalpanels = CDbl(DesignTotalPanels())
    totalarraywatts = panelwatts * totalpanels
    
    LiveArrayKw = CDbl(totalarraywatts / 1000)
End Function

Public Function LiveStringVmp() As Double
    Dim stringtotal As Double
    Dim vmp As Double
    Dim totalstringvmp As Double
    
    stringtotal = CDbl(PanelsPerString())
    vmp = CDbl(EquipmentPanelVmp())
    totalstringvmp = stringtotal * vmp
    
    LiveStringVmp = CDbl(totalstringvmp)
End Function

Public Function TestColdStringVoc()
    MsgBox "ColdStringVoc:" & vbCrLf & ColdStringVoc(), vbInformation, "ColdStringVoc Test"
End Function
