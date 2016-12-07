<!-- #include file=_core.asp -->
<!-- #include file=_chart2.asp -->
<%
'****************************************************************
' Filename..: clientsummary.asp
' Date......: 11/30/2016
' Purpose...: client agent deployment summary
'****************************************************************
time1 = Timer

PageTitle = "Computers by Client Status"
SortBy  = CMWT_GET("s", "Caption0")
QueryON = CMWT_GET("qq", "")

tcount = CMWT_CM_CLIENTCOUNT()

CMWT_NewPage "", "", ""
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT CASE WHEN Client0 = 1 THEN 'Client Installed' " & _
	"ELSE 'No Client' END AS ItemName, COUNT(DISTINCT ResourceID) AS QTY " & _
	"FROM dbo.v_R_System " & _
	"GROUP BY Client0 ORDER BY ItemName"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>
	
</body>
</html>