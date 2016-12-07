<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmtasks.asp
' lastupdate..... 11/30/2016
' description.... site maintenance tasks report
'-----------------------------------------------------------------------------
time1 = Timer

SortBy    = CMWT_GET("s", "TaskName")
QueryOn   = CMWT_GET("qq", "")
PageTitle = "Maintenance Tasks"

CMWT_NewPage "", "", ""
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT T1.TaskName, " & _
	"CASE WHEN T1.VX LIKE '%IsEnabled=_1_ %' THEN 'YES' ELSE 'NO' END AS Enabled " & _
	"FROM " & _
	"(SELECT ItemName AS TaskName, " & _
		"CONVERT(VARCHAR(255),[Value]) AS VX " & _
		"FROM dbo.SC_MISCItem) " & _
	"AS T1 " & _
	"ORDER BY " & SortBy

	Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmtasks.asp", "TASKNAME=cmtask.asp?tn="
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>