<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlreports.asp
' lastupdate..... 12/07/2016
' description.... custom reports list
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "ReportName")
QueryON = CMWT_GET("qq", "")

PageTitle    = "SQL Reports"
PageBackLink = "reports.asp"
PageBackName = "Reports"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT " & _
	"ReportID, ReportType, ReportName, " & _
	"CreatedBy, DateCreated, Comment " & _
	"FROM dbo.Reports2 " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
CMWT_DB_TABLEGRID rs, "", "customreports.asp", ""
CMWT_DB_CLOSE()
Response.Write "<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h30"" value=""New Report"" onClick=""document.location.href='sqlrepadd.asp'"" /><br/>"
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
