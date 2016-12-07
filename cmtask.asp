<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: cmtask.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: site maintenance task details
'****************************************************************
time1 = Timer

TaskName = CMWT_GET("tn", "")
SortBy   = CMWT_GET("s", "TaskName")
QueryON  = CMWT_GET("qq", "")
PageTitle = "Maintenance Task"

CMWT_NewPage "", "", ""
PageBackLink = "cmtasks.asp"
PageBackName = "Maintenance Tasks"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%


query = "SELECT [TaskName]," & _
	"[TaskType],[IsEnabled],[NumRefreshDays],[DaysOfWeek]," & _
	"[BeginTime],[LatestBeginTime],[BackupLocation],[DeleteOlderThan] " & _
	"FROM [dbo].[vSMS_SC_SQL_Task] " & _
	"WHERE TaskName='" & TaskName & "'"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>