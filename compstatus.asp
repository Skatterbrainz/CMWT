<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: compstatus.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: site component status summary
'****************************************************************
time1 = Timer

PageTitle = "Component Status"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "ComponentName")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"a.ComponentName, " & _
	"Infos, " & _
	"Warnings, " & _
	"Errors, " & _
	"LastContacted " & _
"FROM " & _
	"dbo.v_ComponentSummarizer a " & _
"INNER JOIN " & _
"(SELECT DISTINCT " & _
	"ComponentName, MAX(LastContacted) as MT " & _
"FROM " & _
	"dbo.v_ComponentSummarizer " & _
"GROUP BY " & _
	"ComponentName) b " & _
"ON a.ComponentName = b.ComponentName " & _
"AND a.LastContacted = b.MT " & _
"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "compstatus.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_Query()
CMWT_Footer()
%>

</body>
</html>