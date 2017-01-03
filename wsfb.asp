<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... wsfb.asp
' lastupdate..... 01/02/2017
' description.... windows store for business configurations
'-----------------------------------------------------------------------------
time1 = Timer

SortBy   = CMWT_GET("s", "TenantID")
QueryON  = CMWT_GET("qq", "")

PageTitle    = "Windows Store for Business"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT TOP 1 " & _
	"TenantID, " & _
	"ClientID, " & _
	"ContentLocation, " & _
	"DefaultLocale, " & _
	"LastSyncStatus, " & _
	"LastSyncTime, " & _
	"LastSuccessfulSyncTime " & _
	"FROM dbo.vWSfBConfigurationData "

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEROWGRID rs, "", "wsfb.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
