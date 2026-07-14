Attribute VB_Name = "modEngineInit"
Option Explicit

Public Const SDEP_SUCCESS As Long = 0
Public Const SDEP_BAD_ID As Long = 1
Public Const SDEP_BUFFER_TOO_SMALL As Long = 2
Public Const SDEP_INVALID_ARGUMENT As Long = 3
Public Const SDEP_DATABASE_NOT_LOADED As Long = 4
Public Const SDEP_DATABASE_CORRUPT As Long = 5
Public Const SDEP_WRONG_DATABASE_VERSION As Long = 6
Public Const SDEP_INTERNAL_ERROR As Long = 1000

Public Const SDEP_STRING_BUFFER_LEN As Long = 512

Private mEngineInitialized As Boolean

#If VBA7 Then
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" _
        (ByVal lpPathName As String) As Long

    Public Declare PtrSafe Function sdep_init Lib "sdep_engine.dll" () As Long
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

Public Sub ResetEngineInitializationFlag()
    mEngineInitialized = False
End Sub

Public Sub EnsureEngineInitialized()
    If Not mEngineInitialized Then InitializeEngine
End Sub

Public Sub RaiseIfFailed(ByVal status As Long, ByVal sourceName As String)
    If status <> SDEP_SUCCESS Then
        Err.Raise vbObjectError + status, sourceName, _
            "SDEP engine call failed. Status code: " & CStr(status)
    End If
End Sub

Public Function TrimNullTerminatedString(ByVal s As String) As String
    Dim zeroPos As Long
    zeroPos = InStr(1, s, vbNullChar)
    If zeroPos > 0 Then
        TrimNullTerminatedString = Left$(s, zeroPos - 1)
    Else
        TrimNullTerminatedString = s
    End If
End Function

Public Function ReadUtf16String(ByVal status As Long, ByVal buf As String, ByVal sourceName As String) As String
    RaiseIfFailed status, sourceName
    ReadUtf16String = TrimNullTerminatedString(buf)
End Function
