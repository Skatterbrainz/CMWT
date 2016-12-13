<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlreports.asp
' lastupdate..... 12/12/2016
' description.... custom reports list
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "ReportName")
QueryON = CMWT_GET("qq", "")

PageTitle    = "Custom Reports"
PageBackLink = "reports.asp"
PageBackName = "Reports"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT DISTINCT " & _
	"ReportID AS RepID, CASE WHEN ReportType=1 THEN 'SQL' WHEN ReportType=2 THEN 'ADDS' ELSE 'OTHER' END AS ReportType, ReportName, " & _
	"CreatedBy, DateCreated, Comment " & _
	"FROM dbo.Reports2 " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
CMWT_DB_TABLEGRID rs, "", "customreports.asp", ""
CMWT_DB_CLOSE()
Response.Write "<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h30"" value=""New Report"" onClick=""document.location.href='sqlrepadd.asp'"" />" & _
	"<input type=""button"" name=""b2"" id=""b2"" class=""btx w150 h30"" title=""Export Reports"" value=""Export"" onClick=""document.location.href='sqlrepexp.asp'"" /><br/>"
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
