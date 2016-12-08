<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmscopes.asp
' lastupdate..... 11/30/2016
' description.... security scopes
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Site Security Scopes"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "ScopeName")
QueryOn = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<% 
	
query = "SELECT DISTINCT CategoryName AS ScopeName, " & _
	"CategoryDescription AS Description,SourceSite AS SiteCode, " & _
	"NumberOfAdmins AS Admins,NumberOfObjects AS Objects " & _
	"FROM dbo.vRBAC_SecuredCategories " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmscopes.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>