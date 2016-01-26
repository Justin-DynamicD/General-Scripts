'on Error Resume Next

set WinScript = createObject("wscript.shell")
set fso = CreateObject("scripting.FileSystemObject")
strPath = InputBox("UNC Path to dept directory:",,"C:\Users test\")

Set Folder = fso.GetFolder(strPath)
Set SubFldrs = Folder.SubFolders
If SubFldrs.Count <> 0 Then
	For Each SubFo in SubFldrs	
		strSubFo = SubFo.Name
		wscript.echo strSubfo
		Set oExec = winScript.exec("xcacls.exe """ & strPath & "\" & strSubfo & "\Scans "" /t /e /c /y /g XeroxScans:F")
		Do While Not oExec.StdOut.AtEndOfStream 
			Readeverything = oExec.stdOut.ReadAll
			wscript.echo ReadEverything
		Loop
	Next
End If
