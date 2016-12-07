<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... models.asp
' lastupdate..... 11/30/2016
' description.... distinct hardware models in hw inventory
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle = "Computer Models"
PageBackLink = "assets.asp"
PageBackName = "Assets"

SortBy  = CMWT_GET("s", "Model0")
QueryOn = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT Model0 AS ModelName, COUNT(*) AS QTY " & _
	"FROM (" & q_devices & ") AS T1 " & _
	"WHERE (T1.Model0 IS NOT NULL) " & _
	"GROUP BY T1.Model0 " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "models.asp", ""
CMWT_DB_CLOSE()

CMWT_SHOW_QUERY() 
CMWT_Footer()
%>
	
</body>
</html>