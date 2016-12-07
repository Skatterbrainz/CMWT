<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... customreports.asp
' lastupdate..... 12/03/2016
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
	
query = "SELECT DISTINCT ReportID,ReportName,Comment,DateCreated " & _
	"FROM dbo.Reports " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
CMWT_DB_TABLEGRID rs, "", "customreports.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()
%>

</body>
</html>