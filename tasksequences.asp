<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... tasksequences.asp
' lastupdate..... 12/29/2016
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

query = "SELECT ALL SMS_TASKSEQUENCEPACKAGE.NAME, " & _
	"SMS_TASKSEQUENCEPACKAGE.PKGID AS TSPKGID, " & _
	"SMS_TASKSEQUENCEPACKAGE.DESCRIPTION, " & _
	"SMS_TASKSEQUENCEPACKAGE.SOURCEDATE, " & _
	"SMS_TASKSEQUENCEPACKAGE.SOURCESITE, " & _
	"CASE WHEN SMS_TASKSEQUENCEPACKAGE.TS_TYPE=1 THEN 'GENERIC' ELSE 'OSD' END AS TS_TYPE " & _
	"FROM VSMS_TASKSEQUENCEPACKAGE AS SMS_TASKSEQUENCEPACKAGE " & _
	"ORDER BY " & SortBy	

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "tasksequences.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
