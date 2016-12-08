<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... siteconfig.asp
' lastupdate..... 11/30/2016
' description.... site configuration table report
'-----------------------------------------------------------------------------
time1 = Timer
SortBy   = CMWT_GET("s", "RoleName")
KeyValue = CMWT_GET("id", "")
KeySet   = CMWT_GET("ks", "1")

PageTitle = "Site Configuration"

CMWT_NewPage "", "", ""
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%


query = "SELECT RoleName,SiteCode,RoleID,State,Configuration," & _
	"MessageID,LastEvaluatingTime,Param1,Param2,Param3,Param4,Param5,Param6 " & _
	"FROM dbo.vCM_SiteConfiguration " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "app.asp?pn=" & pn, ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>

</html>