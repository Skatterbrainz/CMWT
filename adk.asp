<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adk.asp
' Author....: David M. Stein
' Date......: 03/21/2016
' Purpose...: windows adk installation information
'****************************************************************
time1 = Timer

pageTitle = "ADK Properties"
SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DeploymentKitVersion AS ADKVersion,NetBiosName,FQDN " & _
	"FROM dbo.vSMS_OSDeploymentKitInstalled"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "adk.asp", ""
CMWT_DB_CLOSE()

query = "SELECT DISTINCT DeploymentKitVersion,ProductType,Name " & _
	"FROM dbo.vSMS_OSDeploymentKitSupportedPlatforms " & _
	"ORDER BY " & SortBy

CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "adk.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_Query()
CMWT_Footer()
%>

</body>
</html>