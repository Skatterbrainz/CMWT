<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmaccounts.asp
' lastupdate..... 12/10/2016
' description.... site accounts
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "FeatureName")
QueryOn = CMWT_GET("qq", "")

PageTitle    = "Site Accounts"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT UsageName AS FeatureName,UserName " & _
	"FROM dbo.vSMS_SC_AccountUsage " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmaccounts.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>
