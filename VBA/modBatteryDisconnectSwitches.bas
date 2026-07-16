Attribute VB_Name = "modBatteryDisconnectSwitches"
Option Explicit

Private Const DEFAULT_BUFFER_LENGTH As Long = 65535

#If VBA7 Then
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_ballpark_device Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_continuous_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_device_type Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_enclosure_ingress Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_id Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_intermittent_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_interrupt_rating_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_manufacturer Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_model Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_overcurrent_protection Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_poles Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_rated_dc_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_remote_operated Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_source_url Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_source_notes Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_battery_disconnect_switches_ul_listed Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#Else
    Private Declare Function sdep_get_battery_disconnect_switches_ids Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_model_list Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_ballpark_device Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_continuous_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_device_type Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_enclosure_ingress Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_id Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_intermittent_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_interrupt_rating_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_manufacturer Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_model Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_overcurrent_protection Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_poles Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_rated_dc_v Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_remote_operated Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_source_url Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_source_notes Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_battery_disconnect_switches_ul_listed Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
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

Public Function GetBatteryDisconnectSwitchesIds(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesIds = sdep_get_battery_disconnect_switches_ids(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesModelList(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesModelList = sdep_get_battery_disconnect_switches_model_list(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesBallparkDevice(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesBallparkDevice = sdep_get_battery_disconnect_switches_ballpark_device(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesContinuousA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesContinuousA = sdep_get_battery_disconnect_switches_continuous_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesDeviceType(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesDeviceType = sdep_get_battery_disconnect_switches_device_type(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesEnclosureIngress(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesEnclosureIngress = sdep_get_battery_disconnect_switches_enclosure_ingress(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesId(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesId = sdep_get_battery_disconnect_switches_id(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesIntermittentA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesIntermittentA = sdep_get_battery_disconnect_switches_intermittent_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesInterruptRatingA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesInterruptRatingA = sdep_get_battery_disconnect_switches_interrupt_rating_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesManufacturer(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesManufacturer = sdep_get_battery_disconnect_switches_manufacturer(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesModel(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesModel = sdep_get_battery_disconnect_switches_model(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesOvercurrentProtection(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesOvercurrentProtection = sdep_get_battery_disconnect_switches_overcurrent_protection(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesPoles(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesPoles = sdep_get_battery_disconnect_switches_poles(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesRatedDcV(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetBatteryDisconnectSwitchesRatedDcV = sdep_get_battery_disconnect_switches_rated_dc_v(VarPtr(idBytes(0)), value)
End Function

Public Function GetBatteryDisconnectSwitchesRemoteOperated(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesRemoteOperated = sdep_get_battery_disconnect_switches_remote_operated(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesSourceUrl(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesSourceUrl = sdep_get_battery_disconnect_switches_source_url(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesSourceNotes(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesSourceNotes = sdep_get_battery_disconnect_switches_source_notes(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetBatteryDisconnectSwitchesUlListed(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetBatteryDisconnectSwitchesUlListed = sdep_get_battery_disconnect_switches_ul_listed(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Sub TestBatteryDisconnectSwitches()
    Dim status As Long
    Dim ids As String
    Dim models As String
    Dim manufacturer As String
    Dim model As String
    Dim numeric_value As Double

    EnsureDllInitialized

    status = GetBatteryDisconnectSwitchesIds(ids)
    RaiseIfFailed status, "TestBatteryDisconnectSwitches"
    status = GetBatteryDisconnectSwitchesModelList(models)
    RaiseIfFailed status, "TestBatteryDisconnectSwitches"
    status = GetBatteryDisconnectSwitchesManufacturer("K000001", manufacturer)
    RaiseIfFailed status, "TestBatteryDisconnectSwitches"
    status = GetBatteryDisconnectSwitchesModel("K000001", model)
    RaiseIfFailed status, "TestBatteryDisconnectSwitches"
    status = GetBatteryDisconnectSwitchesBallparkDevice("K000001", numeric_value)
    RaiseIfFailed status, "TestBatteryDisconnectSwitches"
    MsgBox _
        "IDs: " & ids & vbCrLf & _
        "Models: " & models & vbCrLf & _
        "Manufacturer: " & manufacturer & vbCrLf & _
        "Model: " & model & vbCrLf & _
        "Numeric value: " & CStr(numeric_value), vbInformation, "SDEP BatteryDisconnectSwitches Test"
End Sub

