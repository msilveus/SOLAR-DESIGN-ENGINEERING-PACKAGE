Attribute VB_Name = "modEngine_utf16_patch"
Option Explicit

' Patch pattern for SDEP UTF-16 string readers.
' Import into a temporary module or copy these sections into modEngine.

Public Const SDEP_SUCCESS As Long = 0
Public Const SDEP_BAD_ID As Long = 1
Public Const SDEP_BUFFER_TOO_SMALL As Long = 2
Public Const SDEP_INVALID_ARGUMENT As Long = 3
Public Const SDEP_DATABASE_NOT_LOADED As Long = 4
Public Const SDEP_DATABASE_CORRUPT As Long = 5
Public Const SDEP_WRONG_DATABASE_VERSION As Long = 6
Public Const SDEP_INTERNAL_ERROR As Long = 1000

Private Const SDEP_STRING_BUFFER_LEN As Long = 512
Private mEngineInitialized As Boolean

#If VBA7 Then
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" _
        (ByVal lpPathName As String) As Long

    Public Declare PtrSafe Function sdep_init Lib "sdep_engine.dll" () As Long

    ' UTF-16 output string functions: buffer is StrPtr(buf), buffer_len is Len(buf)
    Public Declare PtrSafe Function sdep_get_panels_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_panels_source_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_inverters_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_inverters_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long

    Public Declare PtrSafe Function sdep_get_batteries_manufacturer Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_model Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_chemistry Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_closed_loop_with_sol_ark Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_comm_interface Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_notes Lib "sdep_engine.dll" _
        (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#End If

Public Sub InitializeEngine()
    Dim dllFolder As String
    Dim status As Long

    If mEngineInitialized Then Exit Sub

    dllFolder = ThisWorkbook.Path

    If Len(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1000, "InitializeEngine", _
            "Workbook must be saved before loading sdep_engine.dll."
    End If

    If Dir(dllFolder & "\sdep_engine.dll") = "" Then
        Err.Raise vbObjectError + 1001, "InitializeEngine", _
            "sdep_engine.dll not found in workbook folder: " & dllFolder
    End If

    If SetDllDirectory(dllFolder) = 0 Then
        Err.Raise vbObjectError + 1002, "InitializeEngine", _
            "Failed to set DLL search directory: " & dllFolder
    End If

    status = sdep_init()
    RaiseIfFailed status, "InitializeEngine"

    mEngineInitialized = True
End Sub

Private Sub EnsureEngineInitialized()
    If Not mEngineInitialized Then InitializeEngine
End Sub

Private Sub RaiseIfFailed(ByVal status As Long, ByVal sourceName As String)
    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, sourceName, _
            "SDEP engine call failed. Status code: " & CStr(status)
    End If
End Sub

Private Function TrimNullTerminatedString(ByVal s As String) As String
    Dim zeroPos As Long
    zeroPos = InStr(1, s, vbNullChar)

    If zeroPos > 0 Then
        TrimNullTerminatedString = Left$(s, zeroPos - 1)
    Else
        TrimNullTerminatedString = s
    End If
End Function

Private Function ReadUtf16String(ByVal status As Long, ByVal buf As String, ByVal sourceName As String) As String
    RaiseIfFailed status, sourceName
    ReadUtf16String = TrimNullTerminatedString(buf)
End Function

' Generic helper for all UTF-16 string calls.
Private Function ReadSdepString(ByVal id As String, ByVal sourceName As String, ByVal selector As Long) As String
    Dim buf As String
    Dim status As Long

    EnsureEngineInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)

    Select Case selector
        Case 1: status = sdep_get_panels_manufacturer(id, StrPtr(buf), Len(buf))
        Case 2: status = sdep_get_panels_model(id, StrPtr(buf), Len(buf))
        Case 3: status = sdep_get_panels_source_notes(id, StrPtr(buf), Len(buf))
        Case 4: status = sdep_get_inverters_manufacturer(id, StrPtr(buf), Len(buf))
        Case 5: status = sdep_get_inverters_model(id, StrPtr(buf), Len(buf))
        Case 6: status = sdep_get_inverters_notes(id, StrPtr(buf), Len(buf))
        Case 7: status = sdep_get_batteries_manufacturer(id, StrPtr(buf), Len(buf))
        Case 8: status = sdep_get_batteries_model(id, StrPtr(buf), Len(buf))
        Case 9: status = sdep_get_batteries_chemistry(id, StrPtr(buf), Len(buf))
        Case 10: status = sdep_get_batteries_closed_loop_with_sol_ark(id, StrPtr(buf), Len(buf))
        Case 11: status = sdep_get_batteries_comm_interface(id, StrPtr(buf), Len(buf))
        Case 12: status = sdep_get_batteries_notes(id, StrPtr(buf), Len(buf))
        Case Else
            Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid string selector."
    End Select

    ReadSdepString = ReadUtf16String(status, buf, sourceName)
End Function

' Public wrappers.
Public Function PanelManufacturer(ByVal PanelID As String) As String
    PanelManufacturer = ReadSdepString(PanelID, "PanelManufacturer", 1)
End Function

Public Function PanelModel(ByVal PanelID As String) As String
    PanelModel = ReadSdepString(PanelID, "PanelModel", 2)
End Function

Public Function PanelSourceNotes(ByVal PanelID As String) As String
    PanelSourceNotes = ReadSdepString(PanelID, "PanelSourceNotes", 3)
End Function

Public Function InverterManufacturer(ByVal InverterID As String) As String
    InverterManufacturer = ReadSdepString(InverterID, "InverterManufacturer", 4)
End Function

Public Function InverterModel(ByVal InverterID As String) As String
    InverterModel = ReadSdepString(InverterID, "InverterModel", 5)
End Function

Public Function InverterNotes(ByVal InverterID As String) As String
    InverterNotes = ReadSdepString(InverterID, "InverterNotes", 6)
End Function

Public Function BatteryManufacturer(ByVal BatteryID As String) As String
    BatteryManufacturer = ReadSdepString(BatteryID, "BatteryManufacturer", 7)
End Function

Public Function BatteryModel(ByVal BatteryID As String) As String
    BatteryModel = ReadSdepString(BatteryID, "BatteryModel", 8)
End Function

Public Function BatteryChemistry(ByVal BatteryID As String) As String
    BatteryChemistry = ReadSdepString(BatteryID, "BatteryChemistry", 9)
End Function

Public Function BatteryClosedLoopWithSolArk(ByVal BatteryID As String) As String
    BatteryClosedLoopWithSolArk = ReadSdepString(BatteryID, "BatteryClosedLoopWithSolArk", 10)
End Function

Public Function BatteryCommInterface(ByVal BatteryID As String) As String
    BatteryCommInterface = ReadSdepString(BatteryID, "BatteryCommInterface", 11)
End Function

Public Function BatteryNotes(ByVal BatteryID As String) As String
    BatteryNotes = ReadSdepString(BatteryID, "BatteryNotes", 12)
End Function
