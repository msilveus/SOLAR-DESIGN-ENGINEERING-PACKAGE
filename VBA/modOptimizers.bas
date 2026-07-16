Attribute VB_Name = "modOptimizers"
Option Explicit

Private Const DEFAULT_BUFFER_LENGTH As Long = 65535

#If VBA7 Then
    Private Declare PtrSafe Function sdep_get_optimizers_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_ballpark_device Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_communication Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_compatible_inverters Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_connector Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_id Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_manufacturer Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_imp_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_input_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_isc_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_module_w Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_output_a Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_max_output_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_min_input_v Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_model Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_modules_device Lib "sdep_engine.dll" (ByVal id As LongPtr, ByRef out_value As Double) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_monitoring Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_optimization Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_optimizer_type Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_peak_efficiency Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_rapid_shutdown Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_required_controller Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_source_url Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_source_notes Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_verification_status Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Private Declare PtrSafe Function sdep_get_optimizers_warranty_years Lib "sdep_engine.dll" (ByVal id As LongPtr, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#Else
    Private Declare Function sdep_get_optimizers_ids Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_model_list Lib "sdep_engine.dll" (ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_ballpark_device Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_communication Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_compatible_inverters Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_connector Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_id Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_manufacturer Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_max_imp_a Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_max_input_v Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_max_isc_a Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_max_module_w Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_optimizers_max_output_a Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_max_output_v Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_min_input_v Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_model Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_modules_device Lib "sdep_engine.dll" (ByVal id As Long, ByRef out_value As Double) As Long
    Private Declare Function sdep_get_optimizers_monitoring Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_optimization Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_optimizer_type Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_peak_efficiency Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_rapid_shutdown Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_required_controller Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_source_url Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_source_notes Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_verification_status Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
    Private Declare Function sdep_get_optimizers_warranty_years Lib "sdep_engine.dll" (ByVal id As Long, ByVal buffer As Long, ByVal buffer_len As Long) As Long
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

Public Function GetOptimizersIds(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersIds = sdep_get_optimizers_ids(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersModelList(ByRef value As String) As Long
    Dim buffer As String
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersModelList = sdep_get_optimizers_model_list(StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersBallparkDevice(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersBallparkDevice = sdep_get_optimizers_ballpark_device(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersCommunication(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersCommunication = sdep_get_optimizers_communication(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersCompatibleInverters(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersCompatibleInverters = sdep_get_optimizers_compatible_inverters(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersConnector(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersConnector = sdep_get_optimizers_connector(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersId(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersId = sdep_get_optimizers_id(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersManufacturer(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersManufacturer = sdep_get_optimizers_manufacturer(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMaxImpA(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMaxImpA = sdep_get_optimizers_max_imp_a(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMaxInputV(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMaxInputV = sdep_get_optimizers_max_input_v(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMaxIscA(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMaxIscA = sdep_get_optimizers_max_isc_a(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMaxModuleW(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetOptimizersMaxModuleW = sdep_get_optimizers_max_module_w(VarPtr(idBytes(0)), value)
End Function

Public Function GetOptimizersMaxOutputA(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMaxOutputA = sdep_get_optimizers_max_output_a(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMaxOutputV(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMaxOutputV = sdep_get_optimizers_max_output_v(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersMinInputV(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMinInputV = sdep_get_optimizers_min_input_v(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersModel(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersModel = sdep_get_optimizers_model(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersModulesDevice(ByVal recordId As String, ByRef value As Double) As Long
    Dim idBytes() As Byte
    idBytes = ToCStringBytes(recordId)
    GetOptimizersModulesDevice = sdep_get_optimizers_modules_device(VarPtr(idBytes(0)), value)
End Function

Public Function GetOptimizersMonitoring(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersMonitoring = sdep_get_optimizers_monitoring(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersOptimization(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersOptimization = sdep_get_optimizers_optimization(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersOptimizerType(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersOptimizerType = sdep_get_optimizers_optimizer_type(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersPeakEfficiency(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersPeakEfficiency = sdep_get_optimizers_peak_efficiency(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersRapidShutdown(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersRapidShutdown = sdep_get_optimizers_rapid_shutdown(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersRequiredController(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersRequiredController = sdep_get_optimizers_required_controller(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersSourceUrl(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersSourceUrl = sdep_get_optimizers_source_url(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersSourceNotes(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersSourceNotes = sdep_get_optimizers_source_notes(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersVerificationStatus(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersVerificationStatus = sdep_get_optimizers_verification_status(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Function GetOptimizersWarrantyYears(ByVal recordId As String, ByRef value As String) As Long
    Dim idBytes() As Byte
    Dim buffer As String
    idBytes = ToCStringBytes(recordId)
    buffer = String$(DEFAULT_BUFFER_LENGTH, vbNullChar)
    GetOptimizersWarrantyYears = sdep_get_optimizers_warranty_years(VarPtr(idBytes(0)), StrPtr(buffer), Len(buffer))
    value = TrimNullTerminator(buffer)
End Function

Public Sub TestOptimizers()
    Dim status As Long
    Dim ids As String
    Dim models As String
    Dim manufacturer As String
    Dim model As String
    Dim numeric_value As Double

    EnsureDllInitialized

    status = GetOptimizersIds(ids)
    RaiseIfFailed status, "TestOptimizers"
    status = GetOptimizersModelList(models)
    RaiseIfFailed status, "TestOptimizers"
    status = GetOptimizersManufacturer("O000001", manufacturer)
    RaiseIfFailed status, "TestOptimizers"
    status = GetOptimizersModel("O000001", model)
    RaiseIfFailed status, "TestOptimizers"
    status = GetOptimizersMaxModuleW("O000001", numeric_value)
    RaiseIfFailed status, "TestOptimizers"
    MsgBox _
        "IDs: " & ids & vbCrLf & _
        "Models: " & models & vbCrLf & _
        "Manufacturer: " & manufacturer & vbCrLf & _
        "Model: " & model & vbCrLf & _
        "Numeric value: " & CStr(numeric_value), vbInformation, "SDEP Optimizers Test"
End Sub

