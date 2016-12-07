<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: summarytasks.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: summary tasks detail report
'****************************************************************
time1 = Timer

PageTitle = "Summary Tasks"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "TaskName")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"TaskName,TaskCommand,RunInterval,LastRunDuration," & _
	"LastSuccessfulCompletionTime,LastRunResult," & _
	"CASE WHEN([Enabled]=1) THEN 1 " & _
	"ELSE 0 END AS [IsEnabled] " & _
	"FROM dbo.v_SummaryTasks " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "summarytasks.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>