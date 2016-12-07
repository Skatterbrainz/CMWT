<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmroles.asp
' lastupdate..... 11/30/2016
' description.... security roles
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Site Roles"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "RoleName")
QueryOn = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT RoleName,RoleDescription,NumberOfAdmins AS Members " & _
	"FROM dbo.vRBAC_Roles " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmroles.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>