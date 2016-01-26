'-------------------------------------------------------------------------------------
'
'replacePrinter.vbs
'This will replace old printer mappings with new ones.
'
'-------------------------------------------------------------------------------------

Set WshNetwork = CreateObject("WScript.Network")
Set objNetwork = WScript.CreateObject("WScript.Network")
Set oPrinters = WshNetwork.EnumPrinterConnections
oldDefault = GetDefaultPrinter

'-------------------------------------------------------------------------------------
'
'Printer Replacment List broken down by original and new mapping
'
'-------------------------------------------------------------------------------------

ReplacePrinter "\\newdc\XeroxExec", "\\sharedfp01\XeroxLand"

'-------------------------------------------------------------------------------------
'
'Functions Below this Point

'-------------------------------------------------------------------------------------

Function ReplacePrinter(oldPrinter, newPrinter)
  For i = 0 to oPrinters.Count - 1 Step 2
    if lcase(Trim(oPrinters.Item(i+1))) = lcase(oldPrinter) then
      WshNetwork.AddWindowsPrinterConnection newPrinter
    if lcase(Trim(oldDefault))=lcase(Trim(oldPrinter)) then
      WshNetwork.SetDefaultPrinter newPrinter
    end if          
    If Err <> 0 OR blnError = True Then
      'WScript.Echo Err
    Else   
      WshNetwork.RemovePrinterConnection oldPrinter, true, true
      Exit Function
    End If 'end if err<>0                     
    End If 'end if printer matches current
  Next                      
End Function

Function GetDefaultPrinter()  
  Set oShell = CreateObject("WScript.Shell")  
  sRegVal = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device"  
  sDefault = ""  
  On Error Resume Next 
    sDefault = oShell.RegRead(sRegVal) 
    sDefault = Left(sDefault ,InStr(sDefault, ",") - 1) 
  On Error Goto 0  
  GetDefaultPrinter = sDefault
End Function 
