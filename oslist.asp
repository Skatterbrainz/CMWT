<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... oslist.asp
' lastupdate..... 12/03/2016
' description.... operating systems inventory counds
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "Caption0")
QueryON = CMWT_GET("qq", "")
tcount  = CMWT_CM_CLIENTCOUNT()
PageTitle = "Operating Systems Summary"
PageBackLink = "software.asp"
PageBackName = "Software"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT COALESCE(Caption0, 'UNKNOWN') AS OSCaption, COUNT(DISTINCT Name0) AS QTY " & _
	"FROM (" & q_devices & ") AS T1 " & _
	"GROUP BY T1.Caption0 " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "oslist.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

%>
</body>
</html>