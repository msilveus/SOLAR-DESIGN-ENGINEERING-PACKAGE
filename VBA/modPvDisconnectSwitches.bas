Attribute VB_Name = "modPvDisconnectSwitches"
Option Explicit

Private Const DEFAULT_BUFFER_LENGTH As Long = 65535

#If VBA7 Then
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_ballpark_switch Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_device_type Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_enclosure Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_fused Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_id Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_lockable_handle Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_manufacturer Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_model Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_outdoor_rated Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_poles Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_rated_continuous_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_rated_dc_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_source_url Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_source_notes Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_suitable_for_pv Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_ul_listed Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_pv_disconnect_switches_visible_blade Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#Else
    Private Declare Function sdep_get_pv_disconnect_switches_ids Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_model_list Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_ballpark_switch Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_device_type Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_enclosure Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_fused Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_id Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_lockable_handle Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_manufacturer Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_model Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_outdoor_rated Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_poles Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_rated_continuous_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_rated_dc_v Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_source_url Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_source_notes Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_suitable_for_pv Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_ul_listed Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_pv_disconnect_switches_visible_blade Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
#End If

Private Function TrimNullTerminator(ByVal text As String) As String
    Dim terminator As Long
    terminator = InStr(1, text, vbNullChar, vbBinaryCompare)
    If terminator > 0 Then
        TrimNullTerminator = Left$(text, terminator - 1)
    Else
        TrimNullTerminator = text
    End If
End Function

Private Function ToCStringBytes(ByVal text As String) As Byte()
    Dim bytes() As Byte
    If Len(text) = 0 Then
        ReDim bytes(0 To 0)
        bytes(0) = 0
    Else
        bytes = StrConv(text, vbFromUnicode)
        ReDim Preserve bytes(0 To UBound(bytes) + 1)
        bytes(UBound(bytes)) = 0
    End If
    ToCStringBytes = bytes
End Function

Public Function GetPvDisconnectSwitchesIds(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesIds = sdep_get_pv_disconnect_switches_ids(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesModelList(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesModelList = sdep_get_pv_disconnect_switches_model_list(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesBallparkSwitch(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetPvDisconnectSwitchesBallparkSwitch = sdep_get_pv_disconnect_switches_ballpark_switch(VarPtr(idBytes(0)), value)
End Function

Public Function GetPvDisconnectSwitchesDeviceType(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesDeviceType = sdep_get_pv_disconnect_switches_device_type(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesEnclosure(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesEnclosure = sdep_get_pv_disconnect_switches_enclosure(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesFused(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesFused = sdep_get_pv_disconnect_switches_fused(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesId(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesId = sdep_get_pv_disconnect_switches_id(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesLockableHandle(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesLockableHandle = sdep_get_pv_disconnect_switches_lockable_handle(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesManufacturer(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesManufacturer = sdep_get_pv_disconnect_switches_manufacturer(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesModel(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesModel = sdep_get_pv_disconnect_switches_model(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesOutdoorRated(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesOutdoorRated = sdep_get_pv_disconnect_switches_outdoor_rated(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesPoles(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetPvDisconnectSwitchesPoles = sdep_get_pv_disconnect_switches_poles(VarPtr(idBytes(0)), value)
End Function

Public Function GetPvDisconnectSwitchesRatedContinuousA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetPvDisconnectSwitchesRatedContinuousA = sdep_get_pv_disconnect_switches_rated_continuous_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetPvDisconnectSwitchesRatedDcV(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetPvDisconnectSwitchesRatedDcV = sdep_get_pv_disconnect_switches_rated_dc_v(VarPtr(idBytes(0)), value)
End Function

Public Function GetPvDisconnectSwitchesSourceUrl(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesSourceUrl = sdep_get_pv_disconnect_switches_source_url(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesSourceNotes(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesSourceNotes = sdep_get_pv_disconnect_switches_source_notes(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesSuitableForPv(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesSuitableForPv = sdep_get_pv_disconnect_switches_suitable_for_pv(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesUlListed(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesUlListed = sdep_get_pv_disconnect_switches_ul_listed(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetPvDisconnectSwitchesVisibleBlade(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetPvDisconnectSwitchesVisibleBlade = sdep_get_pv_disconnect_switches_visible_blade(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Sub TestPvDisconnectSwitches()
    Dim status As Long
    Dim ids As String
    Dim models As String
    Dim manufacturer As String
    Dim model As String
    Dim numeric_value As Double

    EnsureDllInitialized

    status = GetPvDisconnectSwitchesIds(ids)
    RaiseIfFailed status, "TestPvDisconnectSwitches"
    status = GetPvDisconnectSwitchesModelList(models)
    RaiseIfFailed status, "TestPvDisconnectSwitches"
    status = GetPvDisconnectSwitchesManufacturer("D000001", manufacturer)
    RaiseIfFailed status, "TestPvDisconnectSwitches"
    status = GetPvDisconnectSwitchesModel("D000001", model)
    RaiseIfFailed status, "TestPvDisconnectSwitches"
    status = GetPvDisconnectSwitchesBallparkSwitch("D000001", numeric_value)
    RaiseIfFailed status, "TestPvDisconnectSwitches"
    MsgBox _
        "IDs: " & ids & vbCrLf & _
        "Models: " & models & vbCrLf & _
        "Manufacturer: " & manufacturer & vbCrLf & _
        "Model: " & model & vbCrLf & _
        "Numeric value: " & CStr(numeric_value), vbInformation, "SDEP PvDisconnectSwitches Test"
End Sub

