Attribute VB_Name = "modWorkbookCalcs"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_string_vmp Lib "sdep_engine.dll" (ByVal panel_vmp As Double, ByVal panels_per_string As Long) As Double
    Public Declare PtrSafe Function sdep_string_voc Lib "sdep_engine.dll" (ByVal panel_voc As Double, ByVal panels_per_string As Long) As Double
#End If

Public Function EngineStringVmp(ByVal PanelVmpValue As Double, ByVal panelsperstring As Long) As Double
    EnsureDllInitialized
    EngineStringVmp = sdep_string_vmp(PanelVmpValue, panelsperstring)
End Function

Public Function EngineStringVoc(ByVal PanelVocValue As Double, ByVal panelsperstring As Long) As Double
    EnsureDllInitialized
    EngineStringVoc = sdep_string_voc(PanelVocValue, panelsperstring)
End Function

Public Function ColdStringVoc() As Double

    ColdStringVoc = _
        PanelVoc() * _
        DesignPanelsPerString() * _
        (1# + Abs(PanelVocTempCoeff()) / 100# * (25# - DesignMinTemperature()))

End Function

Public Function HotStringVoc() As Double

    HotStringVoc = _
        PanelVoc() * _
        DesignPanelsPerString() * _
        (1# + Abs(PanelVocTempCoeff()) / 100# * (25# - DesignMaxTemperature()))

End Function

Private Function StringVmpAtTemperature(ByVal temperatureC As Double) As Double

    Const STC_TEMP_C As Double = 25#

    Dim coefficientPerC As Double

    ' Database value such as -0.24 means -0.24 percent/°C.
    coefficientPerC = PanelPmaxTempCoeff() / 100#

    StringVmpAtTemperature = _
        PanelVmp() * _
        DesignPanelsPerString() * _
        (1# + coefficientPerC * (temperatureC - STC_TEMP_C))

End Function

Public Function ColdStringVmp() As Double
    ColdStringVmp = StringVmpAtTemperature(DesignMinTemperature())
End Function

Public Function NormalStringVmp() As Double
    NormalStringVmp = StringVmpAtTemperature(DesignOperationTemperature())
End Function

Public Function HotStringVmp() As Double
    HotStringVmp = StringVmpAtTemperature(DesignMaxTemperature())
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

Public Function ElectricalPanelCount() As Long
    ElectricalPanelCount = _
        CLng(DesignPanelsPerString()) * _
        CLng(DesignTotalStrings())
End Function

Public Function UnassignedPanelCount() As Long
    UnassignedPanelCount = modWorkbookLayout.LivePanelQty() - ElectricalPanelCount()
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

Public Function LiveNormalStringVmp() As Double
    Dim stringtotal As Double
    Dim vmp As Double
    Dim totalstringvmp As Double
    
    stringtotal = CDbl(DesignPanelsPerString())
    vmp = CDbl(EquipmentPanelVmp())
    totalstringvmp = stringtotal * vmp
    
    LiveNormalStringVmp = CDbl(totalstringvmp)
End Function

Public Function LiveMpptOpA() As Double
    Dim totalports As Double
    Dim impa As Double
    Dim mpptopa As Double
    
    totalports = CDbl(BuildStringsPerMppt())
    impa = CDbl(EquipmentPanelImp())
    mpptopa = totalports * impa
    
    LiveMpptOpA = CDbl(mpptopa)
End Function

Public Function LiveMpptScA() As Double
    Dim totalports As Double
    Dim isca As Double
    Dim mpptsca As Double
    
    totalports = CDbl(BuildStringsPerMppt())
    isca = CDbl(EquipmentPanelIsc())
    mpptsca = totalports * isca
    
    LiveMpptScA = CDbl(mpptsca)
End Function

Public Function PowerMpptKw() As Double
    Dim stringvmp As Double
    Dim mpptopa As Double
    Dim mpptkw As Double
    
    stringvmp = CDbl(modWorkbookLayout.LiveNormalStringVmp())
    mpptopa = CDbl(modWorkbookLayout.LiveMpptOpA())
    mpptkw = mpptopa * stringvmp
    
    PowerMpptKw = CDbl(mpptkw / 1000)
End Function

Public Function LiveBatteryKw() As Double
    Dim battqty As Double
    Dim battkw As Double
    Dim totalbattkw As Double
    
    battqty = CDbl(DesignBatteryQty())
    battkw = CDbl(BatteryKw())
    totalbattkw = battqty * battkw
    
    LiveBatteryKw = CDbl(totalbattkw)
End Function

Public Function TotalPanelCost() As Double
    Dim panelqty As Double
    Dim panelcost As Double
    Dim totalcost As Double
    
    panelqty = CDbl(DesignTotalPanels())
    panelcost = CDbl(EquipmentPanelCost())
    totalcost = panelqty * panelcost
    
    TotalPanelCost = CDbl(totalcost)
End Function

Public Function TotalBatteryCost() As Double
    Dim qty As Double
    Dim cost As Double
    Dim totalcost As Double
    
    qty = CDbl(DesignBatteryQty())
    cost = CDbl(EquipmentBatteryCost())
    totalcost = qty * cost
    
    TotalBatteryCost = CDbl(totalcost)
End Function

Public Function EstimatedEquipmentCost() As Double
    Dim panelcost As Double
    Dim battcost As Double
    Dim invcost As Double
    Dim boscost As Double
    
    panelcost = CDbl(TotalPanelCost())
    battcost = CDbl(TotalBatteryCost())
    invcost = CDbl(EquipmentInverterCost())
    boscost = CDbl(DesignBosCost())
    
    EstimatedEquipmentCost = CDbl(panelcost + battcost + invcost + boscost)
End Function

Public Function CostPerDcWatt() As Double
    Dim watts As Double
    Dim equipmentcost As Double
    Dim invcost As Double
    Dim boscost As Double
    
    watts = CDbl(LiveDcArrayKw()) * 1000
    equipmentcost = CDbl(EstimatedEquipmentCost())
    
    CostPerDcWatt = CDbl(equipmentcost / watts)
End Function

Public Function ArrayWidth() As Double
    Dim width As Double
    Dim orientation As String
    
    orientation = DesignOrientation()
    
        If (orientation = "Portrait") Then
        width = DesignBankWidth() * EquipmentPanelWidth()
    Else
        width = DesignBankWidth() * EquipmentPanelLength()
    End If
    
    ArrayWidth = CDbl(width)
    
End Function

Public Function ArrayHeight() As Double
    Dim height As Double
    Dim orientation As String
    
    orientation = DesignOrientation()
    
    If (orientation = "Portrait") Then
        height = DesignBankHeight() * EquipmentPanelLength()
    Else
        height = DesignBankHeight() * EquipmentPanelWidth()
    End If
    
    ArrayHeight = CDbl(height)
    
End Function

Public Function GroundDepth() As Double
    Dim depth As Double
    Dim height As Double
    Dim tilt As Double
    
    'F17*COS(RADIANS($B$14))
    height = LiveArraySlopedHeight()
    tilt = DesignBankTilt()
    
    depth = height * Cos(Application.WorksheetFunction.Radians(tilt))
    
    GroundDepth = CDbl(depth)
    
End Function

Public Function RearEdgeRise() As Double
    Dim rise As Double
    Dim height As Double
    Dim tilt As Double
    
    'F17*COS(RADIANS($B$14))
    height = LiveArraySlopedHeight()
    tilt = DesignBankTilt()
    
    rise = height * Sin(Application.WorksheetFunction.Radians(tilt))
    
    RearEdgeRise = CDbl(rise)
    
End Function

'*************************
' Status checks
'*************************

Public Function StringVmpCheck() As DesignStatus
    Dim stringvmp As Double
    Dim mpptmin As Double
    Dim mpptmax As Double
    Dim status As Boolean
    
    stringvmp = modWorkbookLayout.LiveNormalStringVmp()
    mpptmin = modWorkbookLayout.EquipmentInvMpptMin()
    mpptmax = modWorkbookLayout.EquipmentInvMpptMax()
    
    StringVmpCheck = CBool(stringvmp >= mpptmin) And CBool(stringvmp <= mpptmax)
End Function

Public Function ColdVmpCheck() As DesignStatus
    Dim stringvmp As Double
    Dim mpptmax As Double
    Dim status As Boolean
    
    stringvmp = ColdStringVmp()
    mpptmax = modWorkbookLayout.EquipmentInvMpptMax()
    
    ColdVmpCheck = CBool(stringvmp <= mpptmax)
End Function

Public Function HotVmpCheck() As DesignStatus
    Dim stringvmp As Double
    Dim mpptmin As Double
    Dim status As Boolean
    
    stringvmp = HotStringVmp()
    mpptmin = modWorkbookLayout.EquipmentInvMpptMin()
    
    HotVmpCheck = CBool(stringvmp >= mpptmin)
End Function

Public Function ColdStringVocCheck() As DesignStatus
    Dim coldvoc As Double
    Dim vocmax As Double
    
    coldvoc = modWorkbookLayout.LiveColdStringVoc()
    vocmax = modWorkbookLayout.EquipmentInvVocMax()
    
    ColdStringVocCheck = CBool(coldvoc <= vocmax)
End Function

Public Function MpptOpACheck() As DesignStatus
    Dim mpptopa As Double
    Dim maxopa As Double
    
    mpptopa = modWorkbookLayout.LiveMpptOpA()
    maxopa = modWorkbookLayout.EquipmentInvOpAMax()
    
    MpptOpACheck = CBool(mpptopa < maxopa)
End Function

Public Function MpptScACheck() As DesignStatus
    Dim mpptsca As Double
    Dim mpptmin As Double
    Dim maxisc As Double
    
    mpptsca = modWorkbookLayout.LiveMpptScA()
    maxisc = modWorkbookLayout.EquipmentInvIscMax()
    
    MpptScACheck = CBool(mpptsca < maxisc)
End Function

Public Function MpptCountCheck() As DesignStatus
    Dim totalstrings As Double
    Dim stringpermppt As Double
    Dim mpptcount As Double
    Dim check As Double
    
    totalstrings = modWorkbookLayout.DesignTotalStrings()
    stringpermppt = modWorkbookLayout.DesignStringsPerMppt()
    mpptcount = modWorkbookLayout.EquipmentInvMpptCount()
    check = totalstrings / stringpermppt
    
    MpptCountCheck = CBool(check <= mpptcount)
End Function

Public Function PanelCountCheck() As DesignStatus
    Dim totalstrings As Double
    Dim panelsperstring As Double
    Dim arraypanels As Double
    Dim calcpanels As Double
    Dim result As Long
    
    result = DS_PASS
    
    totalstrings = modWorkbookLayout.DesignTotalStrings()
    panelsperstring = modWorkbookLayout.DesignPanelsPerString()
    arraypanels = DesignTotalPanels()
    calcpanels = totalstrings * panelsperstring
    
    If arraypanels > calcpanels Then
        result = DS_WARN
    End If
    
    If arraypanels < calcpanels Then
        result = DS_FAIL
    End If
    
    PanelCountCheck = CLng(result)
    
End Function

Public Function NominalVmpStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not StringVmpCheck() Then
        status = STATUS_FAIL
    End If
    
    NominalVmpStatus = CStr(status)
        
End Function

Private Function MaxStatus( _
        ByVal a As DesignStatus, _
        ByVal b As DesignStatus) As DesignStatus

    If a >= b Then
        MaxStatus = a
    Else
        MaxStatus = b
    End If

End Function

Public Function ElectricalStatus() As String
    Dim status As String
    Dim dsstatus As DesignStatus
    
    dsstatus = DS_PASS
    status = STATUS_PASS
    
    dsstatus = MaxStatus(dsstatus, StringVmpCheck())
    dsstatus = MaxStatus(dsstatus, HotVmpCheck())
    dsstatus = MaxStatus(dsstatus, ColdVmpCheck())
    dsstatus = MaxStatus(dsstatus, ColdStringVocCheck())
    dsstatus = MaxStatus(dsstatus, MpptOpACheck())
    dsstatus = MaxStatus(dsstatus, MpptScACheck())
    dsstatus = MaxStatus(dsstatus, MpptCountCheck())
    dsstatus = MaxStatus(dsstatus, PanelCountCheck())
    
    If dsstatus = DS_WARN Then
        status = STATUS_WARN
    End If
    If dsstatus = DS_FAIL Then
        status = STATUS_FAIL
    End If
    
    
    ElectricalStatus = CStr(status)
        
End Function

Public Function HotVmpStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not HotVmpCheck() Then
        status = STATUS_FAIL
    End If
    
    HotVmpStatus = CStr(status)
        
End Function

Public Function ColdVmpStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not ColdVmpCheck() Then
        status = STATUS_FAIL
    End If
    
    ColdVmpStatus = CStr(status)
        
End Function

Public Function ColdVocStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not ColdStringVocCheck() Then
        status = STATUS_FAIL
    End If
    
    ColdVocStatus = CStr(status)
        
End Function

Public Function MpptOpAStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not MpptOpACheck() Then
        status = STATUS_FAIL
    End If
    
    MpptOpAStatus = CStr(status)
        
End Function

Public Function MpptScAStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not MpptScACheck() Then
        status = STATUS_FAIL
    End If
    
    MpptScAStatus = CStr(status)
        
End Function

Public Function MpptCountStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If Not MpptCountCheck() Then
        status = STATUS_FAIL
    End If
    
    MpptCountStatus = CStr(status)
        
End Function

Public Function PanelCountStatus() As String
    Dim status As String
    
    status = STATUS_PASS
    
    If PanelCountCheck() = DS_WARN Then
        status = STATUS_WARN
    End If
    
    If PanelCountCheck() = DS_FAIL Then
        status = STATUS_FAIL
    End If
    
    PanelCountStatus = CStr(status)
        
End Function

'**********************
' Test functions
'**********************
Public Function TestColdStringVoc()
    MsgBox "ColdStringVoc:" & vbCrLf & ColdStringVoc(), vbInformation, "ColdStringVoc Test"
End Function

Public Function TestVmpTemps()
    MsgBox "Vmp Temps:" & vbCrLf & HotStringVmp() & vbCrLf & LiveNormalStringVmp() & vbCrLf & ColdStringVmp(), vbInformation, "Vmp Temps Test"
End Function

Public Function TestLiveMpptOpA()
    MsgBox "LiveMpptOpA:" & vbCrLf & LiveMpptOpA(), vbInformation, "LiveMpptOpA Test"
End Function

Public Function TestPowerMpptKw()
    MsgBox "PowerMpptKw:" & vbCrLf & PowerMpptKw(), vbInformation, "PowerMpptKw Test"
End Function

Public Function TestLiveBatteryKw()
    MsgBox "LiveBatteryKw:" & vbCrLf & LiveBatteryKw(), vbInformation, "LiveBatteryKw Test"
End Function

Public Function TestLiveEstimatedEquipmentCost()
    MsgBox "EstimatedEquipmentCost:" & vbCrLf & EstimatedEquipmentCost(), vbInformation, "EstimatedEquipmentCost Test"
End Function

Public Function TestOrientation()
    MsgBox "DesignOrientation:" & vbCrLf & DesignOrientation(), vbInformation, "DesignOrientation Test"
End Function

Public Function TestGroundDepth()
    MsgBox "GroundDepth:" & vbCrLf & GroundDepth(), vbInformation, "GroundDepth Test"
End Function

Public Function TestRearEdgeRise()
    MsgBox "RearEdgeRise:" & vbCrLf & RearEdgeRise(), vbInformation, "RearEdgeRise Test"
End Function

Public Function TestElectricalStatus()
    MsgBox "ElectricalStatus:" & vbCrLf & ElectricalStatus(), vbInformation, "ElectricalStatus Test"
End Function

Public Function TestStringVmpCheck()
    MsgBox "StringVmpCheck:" & vbCrLf & StringVmpCheck(), vbInformation, "StringVmpCheck Test"
End Function


