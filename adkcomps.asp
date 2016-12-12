<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adkcomps.asp
' Author....: David M. Stein
' Date......: 12/10/2016
' Purpose...: windows adk components
'****************************************************************
time1 = Timer

PageTitle    = "Boot Image Components"
PageBackLink = "software.asp"
PageBackName = "Software"

SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"DeploymentKitVersion AS ADKVersion, " & _
	"UniqueID,Architecture,ComponentID,Name," & _
	"MsiComponentID,Size " & _
	"FROM dbo.vSMS_OSDeploymentKitWinPEOptionalComponents " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "adkcomps.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_Query()
CMWT_Footer()
Response.Write "</body></html>"
%>
