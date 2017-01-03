<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adr.asp
' lastupdate..... 12/29/2016
' description.... automatic deployment rule
'-----------------------------------------------------------------------------
time1 = Timer

KeyVal  = CMWT_GET("id", "")

CMWT_VALIDATE KeyVal, "Record id was not provided"
QueryON = CMWT_GET("qq", "")

PageTitle    = "ADR"
PageBackLink = "adrs.asp"
PageBackName = "Automatic Deployments"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT " & _ 
	"dbo.vSMS_AutoDeployments.AutoDeploymentID AS ADID, " & _ 
	"dbo.vSMS_AutoDeployments.Name, " & _ 
	"dbo.vSMS_AutoDeployments.Description, " & _ 
	"dbo.vSMS_AutoDeployments.EvaluateRule, " & _ 
    "dbo.vSMS_AutoDeployments.AutoDeploymentEnabled, " & _ 
	"dbo.vSMS_AutoDeployments.LastRunTime, " & _ 
	"dbo.vSMS_AutoDeployments.SecurityKey, " & _ 
	"dbo.vSMS_AutoDeployments.AssociatedUpdateGroupID, " & _ 
    "dbo.vSMS_AutoDeployments.IsServicingPlan, " & _ 
	"dbo.vSMS_AutoDeployments.AssociatedDeploymentID, " & _ 
	"dbo.vSMS_ADRDeploymentSettings.CollectionName, " & _ 
	"dbo.vSMS_ADRDeploymentSettings.CollectionID "  & _ 
	"FROM dbo.vSMS_AutoDeployments INNER JOIN " & _ 
    "dbo.vSMS_ADRDeploymentSettings ON " & _ 
	"dbo.vSMS_AutoDeployments.AutoDeploymentID = dbo.vSMS_ADRDeploymentSettings.RuleID " & _ 
	"WHERE AutoDeploymentID=" & KeyVal
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEROWGRID rs, "", "adr.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
