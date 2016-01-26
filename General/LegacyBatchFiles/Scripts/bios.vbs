On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("BRANDONT")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      strBiosCharacteristics = Join(objItem.BiosCharacteristics, ",")
         WScript.Echo "BiosCharacteristics: " & strBiosCharacteristics
      strBIOSVersion = Join(objItem.BIOSVersion, ",")
         WScript.Echo "BIOSVersion: " & strBIOSVersion
      WScript.Echo "BuildNumber: " & objItem.BuildNumber
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "CodeSet: " & objItem.CodeSet
      WScript.Echo "CurrentLanguage: " & objItem.CurrentLanguage
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "IdentificationCode: " & objItem.IdentificationCode
      WScript.Echo "InstallableLanguages: " & objItem.InstallableLanguages
      WScript.Echo "InstallDate: " & WMIDateStringToDate(objItem.InstallDate)
      WScript.Echo "LanguageEdition: " & objItem.LanguageEdition
      strListOfLanguages = Join(objItem.ListOfLanguages, ",")
         WScript.Echo "ListOfLanguages: " & strListOfLanguages
      WScript.Echo "Manufacturer: " & objItem.Manufacturer
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "OtherTargetOS: " & objItem.OtherTargetOS
      WScript.Echo "PrimaryBIOS: " & objItem.PrimaryBIOS
      WScript.Echo "ReleaseDate: " & WMIDateStringToDate(objItem.ReleaseDate)
      WScript.Echo "SerialNumber: " & objItem.SerialNumber
      WScript.Echo "SMBIOSBIOSVersion: " & objItem.SMBIOSBIOSVersion
      WScript.Echo "SMBIOSMajorVersion: " & objItem.SMBIOSMajorVersion
      WScript.Echo "SMBIOSMinorVersion: " & objItem.SMBIOSMinorVersion
      WScript.Echo "SMBIOSPresent: " & objItem.SMBIOSPresent
      WScript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
      WScript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
      WScript.Echo "Status: " & objItem.Status
      WScript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function