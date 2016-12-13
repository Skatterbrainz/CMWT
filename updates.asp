<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... updates.asp
' lastupdate..... 12/13/2016
' description.... deployment status summary
'-----------------------------------------------------------------------------
time1 = Timer

SortBy    = CMWT_GET("s", "BulletinID")
FilterFN  = CMWT_GET("fn", "")
FilterFV  = CMWT_GET("fv", "")
QueryOn   = CMWT_GET("qq", "")
PageTitle = "Software Updates"
PageBackLink = "software.asp"
PageBackName = "Software"

If FilterFN <> "" And FilterFV <> "" Then
	subselect = "WHERE (" & FilterFN & "='" & FilterFV & "')"
	Filtered = True
	PageTitle = "<a href=""updates.asp"">Software Updates</a>: " & FilterFV
Else
	subselect = ""
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT CI_ID AS UID, BulletinID, ArticleID, SeverityName, " & _
	"NumTotal AS Scanned, NumMissing AS Missing, NumPresent AS Installed, " & _
	"NumNotApplicable AS NotReqd, NumUnknown AS Unknown, PercentCompliant AS Compliant, " & _
	"CASE WHEN IsExpired=1 THEN 'YES' ELSE 'NO' END AS Expired, " & _
	"CASE WHEN IsSuperseded=1 THEN 'YES' ELSE 'NO' END AS Superseded," & _
	"CASE WHEN IsDeployed=1 THEN 'YES' ELSE 'NO' END AS Deployed," & _
	"LastStatusTime " & _
	"FROM dbo.vSMS_SoftwareUpdate " & _
	subselect & " ORDER BY " & SortBy
	
CMWT_DEBUG query

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
'CMWT_DB_TableGrid rs, "", "updates.asp", ""
CMWT_DB_TableGridFilter rs, "", "updates.asp", "", "SeverityName", "updates.asp?fn=X&fv=Y"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

Response.Write "</body></html>"
%>
