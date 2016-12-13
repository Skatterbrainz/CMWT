<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: compstatus.asp
' Author....: David M. Stein
' Date......: 12/13/2016
' Purpose...: site component status summary
'****************************************************************
time1 = Timer

PageTitle    = "Component Status"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy   = CMWT_GET("s", "Component")
FilterFN = CMWT_GET("fn", "")
FilterFV = CMWT_GET("fv", "")
QueryON  = CMWT_GET("qq", "")

If FilterFN = "" Then
	query = "SELECT T1.Status, T1.[Site Code], T1.Component, T1.Errors, T1.Warnings, T1.Information, T1.[Last Status Message] FROM ("
Else
	query = "SELECT T1.* FROM ("
	PageTitle = "<a href=""compstatus.asp"">Component Status</a>: " & FilterFV
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%


query = query & _
	"SELECT distinct " & _
	"Case v_ComponentSummarizer.Status " & _
	"When 0 Then 'OK' " & _
	"When 1 Then 'Warning' " & _
	"When 2 Then 'Critical' " & _
	"Else ' ' " & _
	"End As 'Status', " & _
	"SiteCode 'Site Code', " & _
	"MachineName 'Site System', " & _
	"ComponentName 'Component', " & _
	"Case v_componentSummarizer.State " & _
	"When 0 Then 'Stopped' " & _
	"When 1 Then 'Started' " & _
	"When 2 Then 'Paused' " & _
	"When 3 Then 'Installing' " & _
	"When 4 Then 'Re-Installing' " & _
	"When 5 Then 'De-Installing' " & _
	"Else ' ' " & _
	"END AS 'Thread State', " & _
	"Errors 'Errors', " & _
	"Warnings 'Warnings', " & _
	"Infos 'Information', " & _
	"Case v_componentSummarizer.Type " & _
	"When 0 Then 'Autostarting' " & _
	"When 1 Then 'Scheduled' " & _
	"When 2 Then 'Manual' " & _
	"ELSE ' ' " & _
	"END AS 'Startup Type', " & _
	"CASE AvailabilityState " & _
	"When 0 Then 'Online' " & _
	"When 3 Then 'Offline' " & _
	"ELSE ' ' " & _
	"END AS 'Availability State', " & _
	"NextScheduledTime 'Next Scheduled', " & _
	"LastStarted 'Last Started', " & _
	"LastContacted 'Last Status Message', " & _
	"LastHeartbeat 'Last Heartbeat', " & _
	"HeartbeatInterval 'Heartbeat Interval', " & _
	"ComponentType 'Type' " & _
	"from dbo.v_ComponentSummarizer " & _
	"Where TallyInterval = '0001128000100008') AS T1"
If FilterFN <> "" And FilterFV <> "" Then
	query = query & " WHERE (T1." & FilterFN & "='" & FilterFV & "') "
End If

query = query & " ORDER BY T1." & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
If FilterFN = "" Then
	CMWT_DB_TABLEGRID rs, "", "compstatus.asp", ""
Else
	CMWT_DB_TABLEROWGRID rs, "", "compstatus.asp", ""
End If
CMWT_DB_CLOSE()
CMWT_SHOW_Query()
CMWT_Footer()
%>

</body>
</html>