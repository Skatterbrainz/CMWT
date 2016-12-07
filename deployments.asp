<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... depsummary.asp
' lastupdate..... 11/30/2016
' description.... deployment status summary
'-----------------------------------------------------------------------------
time1 = Timer

SortBy    = CMWT_GET("s", "ApplicationName")
QueryOn   = CMWT_GET("qq", "")
PageTitle = "Deployments"
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"dbo.v_ApplicationAssignment.AssignmentID, " & _
	"dbo.v_ApplicationAssignment.ApplicationName, " & _
	"dbo.v_Package.PackageID, " & _
	"dbo.v_ApplicationAssignment.CollectionID, " & _
	"dbo.v_ApplicationAssignment.CollectionName, " & _
	"CASE " & _
	"WHEN dbo.v_ApplicationAssignment.AssignmentEnabled = 1 THEN 'YES' " & _
	"ELSE 'NO' " & _
	"END AS Enabled, " & _
	"CASE " & _
	"WHEN dbo.v_Collection.CollectionType = 1 THEN 'USER' " & _
	"WHEN dbo.v_Collection.CollectionType = 2 THEN 'DEVICE' " & _
	"END AS CollectionType, " & _
	"dbo.v_Collection.MemberCount " & _
	"FROM " & _
	"dbo.v_ApplicationAssignment INNER JOIN " & _
	"dbo.v_Collection ON dbo.v_ApplicationAssignment.CollectionID = dbo.v_Collection.CollectionID LEFT OUTER JOIN " & _
	"dbo.v_Package ON dbo.v_ApplicationAssignment.ApplicationName = dbo.v_Package.Name " & _
	"ORDER BY " & SortBy
	
CMWT_DEBUG query

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "deployments.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

Response.Write "</body></html>"
%>
