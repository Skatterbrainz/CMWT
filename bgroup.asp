<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... bgroup.asp
' lastupdate..... 11/30/2016
' description.... site boundary group details
'-----------------------------------------------------------------------------
time1 = Timer

BGroup  = CMWT_GET("gn", "")
CMWT_VALIDATE BGroup, "Group Name was not specified"

SortBy  = CMWT_GET("s", "Value")
QueryOn = CMWT_GET("qq", "")
PageTitle = BGroup
PageBackLink = "bgroups.asp"
PageBackName = "Boundary Groups"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT dbo.vSMS_BoundaryGroup.GroupID, dbo.vSMS_BoundaryGroup.Name, " & _
	"dbo.vSMS_BoundaryGroup.Description, dbo.vSMS_Boundary.DisplayName, " & _
	"dbo.vSMS_Boundary.Value " & _
	"FROM dbo.vSMS_BoundaryGroup INNER JOIN " & _
	"dbo.vSMS_BoundaryGroupMembers ON dbo.vSMS_BoundaryGroup.GroupID = dbo.vSMS_BoundaryGroupMembers.GroupID INNER JOIN " & _
	"dbo.vSMS_Boundary ON dbo.vSMS_BoundaryGroupMembers.BoundaryID = dbo.vSMS_Boundary.BoundaryID LEFT OUTER JOIN " & _
	"dbo.vSMS_BoundaryGroupSiteSystems ON dbo.vSMS_BoundaryGroupMembers.GroupID = dbo.vSMS_BoundaryGroupSiteSystems.GroupID " & _
	"WHERE dbo.vSMS_BoundaryGroup.Name='" & BGroup & "' " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "bgroup.asp?gn=" & BGroup, ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>
	
</body>
</html>