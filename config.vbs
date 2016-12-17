'----------------------------------------------------------------------
' filename: config.vbs
' author: David Stein
' date created: 11/30/2016
' date modified: 12/16/2016
'----------------------------------------------------------------------

Dim ScriptPath, objFSO, objFile, strFileContents, InputFile, OutputFile
Dim newContent, newSet, newVal, oldVal

Const Caption = "CMWT Setup"
Const ConfigFile = "_config.txt"

ScriptPath = Replace(wscript.ScriptFullName, "\" & wscript.ScriptName, "")

InputFile  = ScriptPath & "\" & ConfigFile
BackupFile = ScriptPath & "\_config.bak"

Const ForReading = 1
Const ForWriting = 2

Sub Update_File (FileName, ContentData)
	Dim objFile
	Set objFile = objFSO.OpenTextFile (FileName, ForWriting, True, 0)
	objFile.Write(ContentData)
	objFile.Close
	Set objFile = Nothing
End Sub

Function Get_ConfigSetting (KeyName, ContentString)
	Dim strLine, tmp, result : result = ""
	For each strLine in Split(ContentString, vbCRLF)
		If Trim(strLine) <> "" And Left(strLine,1) <> ";" Then
			'Wscript.Echo "LINE: " & strLine
			tmp = Split(strLine,"~")
			If LCase(Trim(tmp(0))) = LCase(KeyName) Then
				result = Trim(tmp(1))
			End If
		End If
	Next
	Get_ConfigSetting = result
End Function

Function Set_NewAssociation (KeyName, NewValue)
	Set_NewAssociation = KeyName & "~" & NewValue
End Function

Set objFSO  = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

Wscript.Echo "reading: " & InputFile
Set objFile = objFSO.OpenTextFile(InputFile, ForReading)

If err.Number = 0 Then
	strFileContents = objFile.ReadAll
	objFile.Close
	
	keys = "CMWT_DOMAIN,CMWT_DOMAINSUFFIX,CMWT_ADMINS,DSN_CMDB," & _
		"DSN_CMWT,CMWT_PhysicalPath,CMWT_SiteServer,CMWT_DomainPath,CMWT_MailServer," & _
		"CMWT_MailSender,CMWT_SupportMail,CMWT_ENABLE_LOGGING,CMWT_MAX_LOG_AGE_DAYS," & _
		"CM_SITECODE,CM_AD_TOOLS,CM_AD_TOOLS_SAFETY," & _
		"CM_AD_TOOLS_ADMINGROUPS,CM_AD_TOOLUSER,CM_AD_TOOLPASS"
	
	newContent = ""
	
	For each key in Split(keys, ",")
		oldVal = Get_ConfigSetting(key, strFileContents)
		newVal = InputBox (Replace(key, "_", " "), Caption, oldVal)
		newSet = Set_NewAssociation(key, newVal)
		If newContent <> "" Then
			newContent = newContent & vbCRLF & newSet
		Else
			newContent = newSet
		End If
	Next
	
	Update_File BackupFile, strFileContents
	Update_File InputFile, newContent
	
	Set objFSO  = Nothing
	MsgBox "Original backed up to: " & BackupFile & vbCRLF & _
		"Updated file: " & InputFile, vbOkOnly+vbInformation, Caption
Else
	Wscript.Echo "Error: Unable to read _config.txt file"
	Set objFSO = Nothing
End If

Wscript.Echo "completed!"