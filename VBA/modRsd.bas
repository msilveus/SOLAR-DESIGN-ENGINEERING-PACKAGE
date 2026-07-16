Attribute VB_Name = "modRsd"
Option Explicit

Private Const DEFAULT_BUFFER_LENGTH As Long = 65535

#If VBA7 Then
    Private Declare PtrSafe Function sdep_get_rsd_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_ballpark_device Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_communication Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_connector Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_function Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_id Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_manufacturer Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_max_imp_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_max_input_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_max_isc_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_max_module_w Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_min_input_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_model Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_modules_device Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_rsd_nec_2017_2020 Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_required_controller Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_source_url Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_source_notes Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_ul_pvrss Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_rsd_warranty_years Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
#Else
    Private Declare Function sdep_get_rsd_ids Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_model_list Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_ballpark_device Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_communication Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_connector Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_function Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_id Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_manufacturer Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_max_imp_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_max_input_v Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_max_isc_a Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_max_module_w Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_min_input_v Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_model Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_modules_device Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_rsd_nec_2017_2020 Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_required_controller Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_source_url Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_source_notes Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_ul_pvrss Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_rsd_warranty_years Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
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

Public Function GetRsdIds(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdIds = sdep_get_rsd_ids(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdModelList(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdModelList = sdep_get_rsd_model_list(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdBallparkDevice(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdBallparkDevice = sdep_get_rsd_ballpark_device(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdCommunication(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdCommunication = sdep_get_rsd_communication(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdConnector(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdConnector = sdep_get_rsd_connector(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdFunction(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdFunction = sdep_get_rsd_function(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdId(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdId = sdep_get_rsd_id(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdManufacturer(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdManufacturer = sdep_get_rsd_manufacturer(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdMaxImpA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdMaxImpA = sdep_get_rsd_max_imp_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdMaxInputV(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdMaxInputV = sdep_get_rsd_max_input_v(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdMaxIscA(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdMaxIscA = sdep_get_rsd_max_isc_a(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdMaxModuleW(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdMaxModuleW = sdep_get_rsd_max_module_w(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdMinInputV(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdMinInputV = sdep_get_rsd_min_input_v(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdModel(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdModel = sdep_get_rsd_model(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdModulesDevice(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdModulesDevice = sdep_get_rsd_modules_device(VarPtr(idBytes(0)), value)
End Function

Public Function GetRsdNec20172020(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdNec20172020 = sdep_get_rsd_nec_2017_2020(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdRequiredController(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdRequiredController = sdep_get_rsd_required_controller(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdSourceUrl(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdSourceUrl = sdep_get_rsd_source_url(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdSourceNotes(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdSourceNotes = sdep_get_rsd_source_notes(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdUlPvrss(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetRsdUlPvrss = sdep_get_rsd_ul_pvrss(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetRsdWarrantyYears(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetRsdWarrantyYears = sdep_get_rsd_warranty_years(VarPtr(idBytes(0)), value)
End Function

Public Sub TestRsd()
    Dim status As Long
    Dim ids As String
    Dim models As String
    Dim manufacturer As String
    Dim model As String
    Dim numeric_value As Double

    EnsureDllInitialized

    status = GetRsdIds(ids)
    RaiseIfFailed status, "TestRsd"
    status = GetRsdModelList(models)
    RaiseIfFailed status, "TestRsd"
    status = GetRsdManufacturer("R000001", manufacturer)
    RaiseIfFailed status, "TestRsd"
    status = GetRsdModel("R000001", model)
    RaiseIfFailed status, "TestRsd"
    status = GetRsdBallparkDevice("R000001", numeric_value)
    RaiseIfFailed status, "TestRsd"
    MsgBox _
        "IDs: " & ids & vbCrLf & _
        "Models: " & models & vbCrLf & _
        "Manufacturer: " & manufacturer & vbCrLf & _
        "Model: " & model & vbCrLf & _
        "Numeric value: " & CStr(numeric_value), vbInformation, "SDEP Rsd Test"
End Sub

