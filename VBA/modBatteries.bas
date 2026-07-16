Attribute VB_Name = "modBatteries"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sdep_get_batteries_ids Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_usable_kwh Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_nominal_v Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_capacity_ah Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_cont_charge_a Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_cont_discharge_a Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_peak_a Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_weight_lb Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_ballpark_battery Lib "sdep_engine.dll" (ByVal id As String, ByRef out_value As Double) As Long
    Public Declare PtrSafe Function sdep_get_batteries_manufacturer Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_model Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_model_list Lib "sdep_engine.dll" (ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_chemistry Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_closed_loop_with_sol_ark Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_comm_interface Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
    Public Declare PtrSafe Function sdep_get_batteries_notes Lib "sdep_engine.dll" (ByVal id As String, ByVal buffer As LongPtr, ByVal buffer_len As LongPtr) As Long
#End If

Private Function BNum(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As Double
    Dim status As Long, value As Double
    EnsureDllInitialized
    Select Case selector
        Case 1: status = sdep_get_batteries_usable_kwh(id, value)
        Case 2: status = sdep_get_batteries_nominal_v(id, value)
        Case 3: status = sdep_get_batteries_capacity_ah(id, value)
        Case 4: status = sdep_get_batteries_cont_charge_a(id, value)
        Case 5: status = sdep_get_batteries_cont_discharge_a(id, value)
        Case 6: status = sdep_get_batteries_peak_a(id, value)
        Case 7: status = sdep_get_batteries_weight_lb(id, value)
        Case 8: status = sdep_get_batteries_ballpark_battery(id, value)
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid battery numeric selector."
    End Select
    RaiseIfFailed status, sourceName
    BNum = value
End Function

Private Function BStr(ByVal id As String, ByVal selector As Long, ByVal sourceName As String) As String
    Dim status As Long, buf As String
    EnsureDllInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    Select Case selector
        Case 1: status = sdep_get_batteries_manufacturer(id, StrPtr(buf), Len(buf))
        Case 2: status = sdep_get_batteries_model(id, StrPtr(buf), Len(buf))
        Case 3: status = sdep_get_batteries_chemistry(id, StrPtr(buf), Len(buf))
        Case 4: status = sdep_get_batteries_closed_loop_with_sol_ark(id, StrPtr(buf), Len(buf))
        Case 5: status = sdep_get_batteries_comm_interface(id, StrPtr(buf), Len(buf))
        Case 6: status = sdep_get_batteries_notes(id, StrPtr(buf), Len(buf))
        Case Else: Err.Raise vbObjectError + SDEP_INVALID_ARGUMENT, sourceName, "Invalid battery string selector."
    End Select
    BStr = ReadUtf16String(status, buf, sourceName)
End Function

Public Function BatteryIDs() As String
    Dim status As Long, buf As String
    EnsureDllInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_ids(StrPtr(buf), Len(buf))
    BatteryIDs = ReadUtf16String(status, buf, "BatteryIDs")
End Function

Public Function BatteryModels() As String
    Dim status As Long, buf As String
    EnsureDllInitialized
    buf = String$(SDEP_STRING_BUFFER_LEN, vbNullChar)
    status = sdep_get_batteries_model_list(StrPtr(buf), Len(buf))
    BatteryModels = ReadUtf16String(status, buf, "BatteryModels")
End Function

Public Function BatteryUsableKwh(ByVal BatteryID As String) As Double: BatteryUsableKwh = BNum(BatteryID, 1, "BatteryUsableKwh"): End Function
Public Function BatteryNominalV(ByVal BatteryID As String) As Double: BatteryNominalV = BNum(BatteryID, 2, "BatteryNominalV"): End Function
Public Function BatteryCapacityAh(ByVal BatteryID As String) As Double: BatteryCapacityAh = BNum(BatteryID, 3, "BatteryCapacityAh"): End Function
Public Function BatteryContChargeA(ByVal BatteryID As String) As Double: BatteryContChargeA = BNum(BatteryID, 4, "BatteryContChargeA"): End Function
Public Function BatteryContDischargeA(ByVal BatteryID As String) As Double: BatteryContDischargeA = BNum(BatteryID, 5, "BatteryContDischargeA"): End Function
Public Function BatteryPeakA(ByVal BatteryID As String) As Double: BatteryPeakA = BNum(BatteryID, 6, "BatteryPeakA"): End Function
Public Function BatteryWeightLb(ByVal BatteryID As String) As Double: BatteryWeightLb = BNum(BatteryID, 7, "BatteryWeightLb"): End Function
Public Function BatteryBallparkCost(ByVal BatteryID As String) As Double: BatteryBallparkCost = BNum(BatteryID, 8, "BatteryBallparkCost"): End Function
Public Function BatteryManufacturer(ByVal BatteryID As String) As String: BatteryManufacturer = BStr(BatteryID, 1, "BatteryManufacturer"): End Function
Public Function BatteryModel(ByVal BatteryID As String) As String: BatteryModel = BStr(BatteryID, 2, "BatteryModel"): End Function
Public Function BatteryChemistry(ByVal BatteryID As String) As String: BatteryChemistry = BStr(BatteryID, 3, "BatteryChemistry"): End Function
Public Function BatteryClosedLoopWithSolArk(ByVal BatteryID As String) As String: BatteryClosedLoopWithSolArk = BStr(BatteryID, 4, "BatteryClosedLoopWithSolArk"): End Function
Public Function BatteryCommInterface(ByVal BatteryID As String) As String: BatteryCommInterface = BStr(BatteryID, 5, "BatteryCommInterface"): End Function
Public Function BatteryNotes(ByVal BatteryID As String) As String: BatteryNotes = BStr(BatteryID, 6, "BatteryNotes"): End Function

Public Sub TestBatteryIDs()
    MsgBox "Battery IDs:" & vbCrLf & BatteryIDs(), vbInformation, "SDEP Battery ID Test"
End Sub

Public Sub TestBatteryModelList()
    MsgBox "Battery Model List:" & vbCrLf & BatteryModels(), vbInformation, "SDEP Battery Model Test"
End Sub

