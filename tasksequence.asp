<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... tasksequence.asp
' lastupdate..... 12/12/2016
' description.... task sequence details
'-----------------------------------------------------------------------------
time1 = Timer

PkgID   = CMWT_GET("id", "")
QueryON = CMWT_GET("qq", "")

CMWT_VALIDATE PkgID, "Package ID was not provided"

PageTitle    = "Task Sequence"
PageBackLink = "tasksequences.asp"
PageBackName = "Task Sequences"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
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
	
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEROWGRID rs, "", "", ""
'CMWT_DB_TABLEGRID rs, "", "tasksequences.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>