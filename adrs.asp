<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adrs.asp
' lastupdate..... 12/29/2016
' description.... automatic deployment rules
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "ADID")
QueryON = CMWT_GET("qq", "")

PageTitle    = "Automatic Deployments"
PageBackLink = "software.asp"
PageBackName = "Software"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT " & _ 
	"dbo.vSMS_AutoDeployments.AutoDeploymentID AS ADID, " & _ 
	"dbo.vSMS_AutoDeployments.Name, " & _ 
	"dbo.vSMS_AutoDeployments.Description, " & _ 
    "dbo.vSMS_AutoDeployments.AutoDeploymentEnabled, " & _ 
    "dbo.vSMS_AutoDeployments.IsServicingPlan, " & _ 
	"dbo.vSMS_AutoDeployments.AssociatedDeploymentID, " & _ 
	"dbo.vSMS_ADRDeploymentSettings.CollectionName, " & _ 
	"dbo.vSMS_ADRDeploymentSettings.CollectionID "  & _ 
	"FROM dbo.vSMS_AutoDeployments INNER JOIN " & _ 
    "dbo.vSMS_ADRDeploymentSettings ON " & _ 
	"dbo.vSMS_AutoDeployments.AutoDeploymentID = dbo.vSMS_ADRDeploymentSettings.RuleID " & _ 
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "adrs.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
