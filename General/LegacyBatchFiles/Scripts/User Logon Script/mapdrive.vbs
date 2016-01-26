'----------------------------------------------------------------------------
'
' mapdrive.vbs
' Drive mapping rules.
'
'----------------------------------------------------------------------------


Option Explicit

Dim objNetwork, objSysInfo, strUserDN
Dim objGroupList, objUser, objFSO
Dim strComputerDN, objComputer
Dim strDFSPath

Set objNetwork = CreateObject("Wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
strUserDN = objSysInfo.userName
strComputerDN = objSysInfo.computerName

' Sets DFS path
strDFSPath = "\\nrpi.local\shared_files\"

' Bind to the user and computer objects with the LDAP provider.
Set objUser = GetObject("LDAP://" & strUserDN)
Set objComputer = GetObject("LDAP://" & strComputerDN)

'----------------------------------------------------------------------------
'
' Map network drives.
'
'----------------------------------------------------------------------------

MapDrive "S:", strDFSPath

'----------------------------------------------------------------------------
'
' Script Cleanup.
'
'----------------------------------------------------------------------------

Set objNetwork = Nothing
Set objFSO = Nothing
Set objSysInfo = Nothing
Set objGroupList = Nothing
Set objUser = Nothing
Set objComputer = Nothing
Set strDFSPath = Nothing

'----------------------------------------------------------------------------
'
'Custom Called Functions Below this Point.  When in doubt, don't touch.
'
'----------------------------------------------------------------------------

Function MapDrive(strDrive, strShare)
' Function to map network share to a drive letter.
' If the drive letter specified is already in use, the function
' attempts to remove the network connection.
' objFSO is the File System Object, with global scope.
' objNetwork is the Network object, with global scope.
' Returns True if drive mapped, False otherwise.

  Dim objDrive

  On Error Resume Next
  If objFSO.DriveExists(strDrive) Then
    Set objDrive = objFSO.GetDrive(strDrive)
    If Err.Number <> 0 Then
      On Error GoTo 0
      MapDrive = False
      Exit Function
    End If
    If CBool(objDrive.DriveType = 3) Then
      objNetwork.RemoveNetworkDrive strDrive, True, True
    Else
      MapDrive = False
      Exit Function
    End If
    Set objDrive = Nothing
  End If
  objNetwork.MapNetworkDrive strDrive, strShare
  If Err.Number = 0 Then
    MapDrive = True
  Else
    Err.Clear
    MapDrive = False
  End If
  On Error GoTo 0
End Function
