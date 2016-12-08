<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmwtlogs.asp
' lastupdate..... 12/03/2016
' description.... cmwt database log maintenance
'-----------------------------------------------------------------------------
time1 = Timer
QueryOn = CMWT_GET("qq", "")

PageTitle    = "CMWT Database Logs"
PageBackLink = "admin.asp"
PageBackName = "Administration"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

if Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
	query1 = "SELECT 'Task records' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.Tasks " & _
		"UNION " & _
		"SELECT 'Task records completed' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.Tasks " & _
		"WHERE DateTimeExecuted IS NOT NULL " & _
		"UNION " & _
		"SELECT 'Task records pending execution' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.Tasks " & _
		"WHERE DateTimeExecuted IS NULL " & _
		"UNION " & _
		"SELECT 'Task records older than " & Application("CMWT_MAX_LOG_AGE_DAYS") & _
		" days' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.Tasks " & _
		"WHERE DateTimeCreated < GETDATE() - " & Application("CMWT_MAX_LOG_AGE_DAYS")

	query2 = "SELECT 'Event log records' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.EventLog " & _
		"UNION " & _
		"SELECT 'Event log records, type INFO' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.EventLog " & _
		"WHERE EventType = 'INFO' " & _
		"UNION " & _
		"SELECT 'Event log records, type ERROR' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.EventLog " & _
		"WHERE EventType = 'ERROR' " & _
		"UNION " & _
		"SELECT 'Event log records older than " & Application("CMWT_MAX_LOG_AGE_DAYS") & _
		" days' AS Description, COUNT(*) AS RECS " & _
		"FROM dbo.EventLog " & _
		"WHERE EventDateTime < GETDATE() - " & Application("CMWT_MAX_LOG_AGE_DAYS")

	query3 = "SELECT DISTINCT EventCategory, COUNT(*) AS QTY " & _
		"FROM dbo.EventLog " & _
		"GROUP BY EventCategory"

	Dim conn, cmd, rs
	CMWT_DB_QUERY Application("DSN_CMWT"), query1
	CMWT_DB_TABLEGRID rs, "", "", ""
	CMWT_DB_CLOSE()

	Response.Write "<br/>"

	CMWT_DB_QUERY Application("DSN_CMWT"), query2
	CMWT_DB_TABLEGRID rs, "", "", ""
	CMWT_DB_CLOSE()

	Response.Write "<br/>"

	CMWT_DB_QUERY Application("DSN_CMWT"), query3
	CMWT_DB_TABLEGRID rs, "", "", ""
	CMWT_DB_CLOSE()
Else
	Response.Write "<table class=""tfx""><tr class=""h200 tr1"">" & _
		"<td class=""td6 ctr v10"">Logging is not enabled.<p>" & _
		"To enable site activity logging, modify the _config.txt file " & _
		"to set CMWT_ENABLE_LOGGING~TRUE</p><p>Then recycle the IIS application pool.</p>" & _
		"</td></tr></table>"
End If

CMWT_Footer()
%>

</body>
</html>