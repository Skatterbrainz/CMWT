<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... tasksequences.asp
' lastupdate..... 12/12/2016
' description.... task sequences report
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

PageTitle    = "Task Sequences"
PageBackLink = "software.asp"
PageBackName = "Software"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT ALL " & _
	"SMS_TaskSequencePackage.Name," & _
	"SMS_TaskSequencePackage.PkgID AS TSPkgID," & _
	"SMS_TaskSequencePackage.Description," & _
	"SMS_TaskSequencePackage.SourceDate, " & _
	"SMS_TaskSequencePackage.SourceSite, " & _
	"CASE " & _
	"WHEN SMS_TaskSequencePackage.TS_Type=1 THEN 'GENERIC' " & _
	"ELSE 'OSD' END AS TS_Type " & _
	"FROM  " & _
		"vSMS_TaskSequencePackage AS SMS_TaskSequencePackage  " & _
	"WHERE SMS_TaskSequencePackage.PkgID NOT IN " & _
		"(SELECT all Folder##Alias##810314.InstanceKey " & _
		"FROM vFolderMembers AS Folder##Alias##810314  " & _
		"WHERE Folder##Alias##810314.ObjectTypeName = N'SMS_TaskSequencePackage') " & _
	"ORDER BY " & SortBy
	

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "tasksequences.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
