<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... tasksequence.asp
' lastupdate..... 12/12/2016
' description.... task sequence details
'-----------------------------------------------------------------------------
time1 = Timer

PkgID   = CMWT_GET("id", "")
PSet    = CMWT_GET("set","1")
QueryON = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s", "ExecutionTime DESC")

CMWT_VALIDATE PkgID, "Package ID was not provided"

PageTitle    = "Task Sequence"
PageBackLink = "tasksequences.asp"
PageBackName = "Task Sequences"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "1=General,2=History"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	mlink = "tasksequence.asp?id=" & PkgID & "&set=" & mset(0)
	If PSet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

Dim conn, cmd, rs

Select Case PSet
	Case "1":
		query = "SELECT ALL " & _
			"SMS_TaskSequencePackage.Name," & _
			"SMS_TaskSequencePackage.PkgID," & _
			"SMS_TaskSequencePackage.Action," & _
			"SMS_TaskSequencePackage.BootImageID," & _
			"SMS_TaskSequencePackage.Category," & _
			"SMS_TaskSequencePackage.CustomHighImpactHeadline," & _
			"SMS_TaskSequencePackage.CustomHighImpactWarning," & _
			"SMS_TaskSequencePackage.CustomProgressMsg," & _
			"SMS_TaskSequencePackage.DependentProgram," & _
			"SMS_TaskSequencePackage.Description," & _
			"SMS_TaskSequencePackage.Duration," & _
			"SMS_TaskSequencePackage.EstimatedDownloadSizeMB," & _
			"SMS_TaskSequencePackage.EstimatedRunTimeMinutes," & _
			"SMS_TaskSequencePackage.DisconnectDelay," & _
			"SMS_TaskSequencePackage.UseForcedDisconnect," & _
			"SMS_TaskSequencePackage.ForcedRetryDelay," & _
			"SMS_TaskSequencePackage.HighImpactTaskSequence," & _
			"SMS_TaskSequencePackage.Icon," & _
			"SMS_TaskSequencePackage.IgnoreSchedule," & _
			"SMS_TaskSequencePackage.ISVString," & _
			"SMS_TaskSequencePackage.Language," & _
			"SMS_TaskSequencePackage.LastRefresh," & _
			"SMS_TaskSequencePackage.LocalizedTaskSequenceDescription," & _
			"SMS_TaskSequencePackage.LocalizedTaskSequenceName," & _
			"SMS_TaskSequencePackage.Manufacturer," & _
			"SMS_TaskSequencePackage.MIFFilename," & _
			"SMS_TaskSequencePackage.MIFName," & _
			"SMS_TaskSequencePackage.MIFPublisher," & _
			"SMS_TaskSequencePackage.MIFVersion," & _
			"SMS_TaskSequencePackage.PackageType, " & _
			"SMS_TaskSequencePackage.PkgFlags," & _
			"SMS_TaskSequencePackage.StorePkgFlag," & _
			"SMS_TaskSequencePackage.PreDownloadRule," & _
			"SMS_TaskSequencePackage.PreferredAddress," & _
			"SMS_TaskSequencePackage.Priority," & _
			"SMS_TaskSequencePackage.ProgramFlags," & _
			"SMS_TaskSequencePackage.ReferencesCount," & _
			"SMS_TaskSequencePackage.RestartRequired," & _
			"SMS_TaskSequencePackage.SedoObjectVersion," & _
			"SMS_TaskSequencePackage.ShareName," & _
			"SMS_TaskSequencePackage.ShareType," & _
			"SMS_TaskSequencePackage.SourceDate," & _
			"SMS_TaskSequencePackage.SourceSite," & _
			"SMS_TaskSequencePackage.SourceVersion," & _
			"SMS_TaskSequencePackage.StoredPkgPath," & _
			"SMS_TaskSequencePackage.StoredPkgVersion," & _
			"SMS_TaskSequencePackage.TS_Flags," & _
			"SMS_TaskSequencePackage.TS_Type," & _
			"SMS_TaskSequencePackage.Version " & _
			"FROM vSMS_TaskSequencePackage AS SMS_TaskSequencePackage " & _
			"WHERE SMS_TaskSequencePackage.PkgID NOT IN " & _
				"(SELECT ALL Folder##Alias##810314.InstanceKey " & _
				"FROM vFolderMembers AS Folder##Alias##810314 " & _
				"WHERE Folder##Alias##810314.ObjectTypeName = N'SMS_TaskSequencePackage') " & _
				"AND SMS_TaskSequencePackage.PkgID = '" & PkgID & "'"
			
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEROWGRID rs, "", "", ""
		'CMWT_DB_TABLEGRID rs, "", "tasksequences.asp", ""
		CMWT_DB_CLOSE()
	Case "2":
		query = "SELECT DISTINCT " & _
			"ExecutionTime,Step,GroupName,ActionName," & _
			"LastStatusMsgID,LastStatusMsgName,ExitCode," & _
			"ActionOutput,ResourceID " & _
			"FROM dbo.vSMS_TaskSequenceExecutionStatus " & _
			"WHERE PackageID='" & PkgID & "' " & _
			"ORDER BY " & SortBY
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEGRID rs, "", "tasksequence.asp?id=" & PkgID & "&set=2", ""
		CMWT_DB_CLOSE()
End Select
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>