<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... depsummary.asp
' lastupdate..... 12/04/2016
' description.... deployment status summary
'-----------------------------------------------------------------------------
time1 = Timer

SortBy    = CMWT_GET("s", "Application")
QueryOn   = CMWT_GET("qq", "")
PageTitle = "Deployment Status Summary"
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"CASE WHEN SoftwareName = '' THEN '(TASK SEQUENCE)' ELSE [SoftwareName] END AS Application," & _
	"CollectionID AS CollID, " & _
	"CollectionName,DeploymentTime AS Deployed, " & _
	"CreationTime AS Created,EnforcementDeadline AS Deadline, " & _
	"NumberSuccess AS Success,NumberInProgress AS InProgress, " & _
	"NumberUnknown AS Unknown,NumberErrors AS Failed, " & _
	"NumberOther AS Other,NumberTotal AS Total, " & _
	"SummarizationTime AS Summarized " & _
	"FROM dbo.v_DeploymentSummary " & _
	"ORDER BY " & SortBy
CMWT_DEBUG query

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "depsummary.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

Response.Write "</body></html>"
%>
