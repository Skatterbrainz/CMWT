<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sitedef.asp
' lastupdate..... 12/10/2016
' description.... site definition settings
'-----------------------------------------------------------------------------
time1 = Timer
QueryOn   = CMWT_GET("qq", "")
DebugOn   = CMWT_GET("debug", "")

PageTitle    = "Site Definition"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT [SiteNumber],[SiteType],[SiteCode],[SiteName]," & _
	"[ParentSiteCode] AS Parent,[SiteServerName] AS Server," & _
	"[SiteServerDomain] AS Domain," & _
	"[SiteServerPlatform] AS CPU,[InstallDirectory] AS InstallPath," & _
	"[SQLServerName] AS SQLHost,[SQLDatabaseName] AS SQL_DB " & _
	"FROM dbo.v_SC_SiteDefinition"
if DebugOn = "1" Then
	response.write "<p>query: " & query & "</p>"
	response.write "<p>" & Application("DSN_CMDB") & "</a>"
	response.end
end If
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
Response.Write "</body></html>"
%>
