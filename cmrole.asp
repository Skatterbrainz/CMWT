<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmrole.asp
' lastupdate..... 11/30/2016
' description.... security role details
'-----------------------------------------------------------------------------
time1 = Timer
RoleName  = CMWT_GET("rn", "")
SortBy    = CMWT_GET("s", "LogonName")
QueryOn   = CMWT_GET("qq", "")
PageTitle = RoleName
PageBackLink = "cmroles.asp"
PageBackName = "Security Roles"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT LogonName, RoleType FROM " & _
	"(SELECT DISTINCT " & _
	"dbo.vRBAC_Roles.RoleName, dbo.vRBAC_Roles.RoleID, " & _
	"dbo.vRBAC_Permissions.LogonName, dbo.vRBAC_Admins.IsGroup, " & _
	"CASE WHEN IsGroup = 'True' THEN 'GROUP' ELSE 'USER' END AS RoleType " & _
	"FROM dbo.vRBAC_Admins " & _
	"INNER JOIN " & _
	"dbo.vRBAC_Permissions ON dbo.vRBAC_Admins.LogonName = dbo.vRBAC_Permissions.LogonName " & _
	"RIGHT OUTER JOIN " & _
	"dbo.vRBAC_Roles ON dbo.vRBAC_Permissions.RoleName = dbo.vRBAC_Roles.RoleName) AS T1 " & _
	"WHERE (RoleName='" & RoleName & "') " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmrole.asp?rn=" & RoleName, ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>