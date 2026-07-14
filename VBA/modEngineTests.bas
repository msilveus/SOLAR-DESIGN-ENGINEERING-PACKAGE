Attribute VB_Name = "modEngineTests"
Option Explicit

Public Sub TestEngineDatabaseReaders()
    Dim msg As String
    InitializeEngine
    msg = "Panel P000001:" & vbCrLf & _
          "  Manufacturer: " & PanelManufacturer("P000001") & vbCrLf & _
          "  Model: " & PanelModel("P000001") & vbCrLf & _
          "  Vmp: " & PanelVmp("P000001") & vbCrLf & _
          "  Voc: " & PanelVoc("P000001") & vbCrLf & vbCrLf & _
          "Inverter I000001:" & vbCrLf & _
          "  Manufacturer: " & InverterManufacturer("I000001") & vbCrLf & _
          "  Model: " & InverterModel("I000001") & vbCrLf & _
          "  MPPT Range: " & InverterMpptMinV("I000001") & " - " & InverterMpptMaxV("I000001") & " V" & vbCrLf & vbCrLf & _
          "Battery B000001:" & vbCrLf & _
          "  Manufacturer: " & BatteryManufacturer("B000001") & vbCrLf & _
          "  Model: " & BatteryModel("B000001") & vbCrLf & _
          "  Chemistry: " & BatteryChemistry("B000001") & vbCrLf & _
          "  Usable kWh: " & BatteryUsableKwh("B000001") & vbCrLf & vbCrLf & _
          "ID Lists:" & vbCrLf & _
          "  Panels: " & PanelIDs() & vbCrLf & _
          "  Inverters: " & InverterIDs() & vbCrLf & _
          "  Batteries: " & BatteryIDs()
    MsgBox msg, vbInformation, "SDEP Engine Database Reader Test"
End Sub
