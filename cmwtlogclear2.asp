<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmwtlogclear2.asp
' lastupdate..... 12/12/2016
' description.... add new custom note
'-----------------------------------------------------------------------------
Response.Expires = -1
KeySet = CMWT_GET("l", "")
RmvSet = CMWT_GET("x1", "")

CMWT_VALIDATE KeySet, "Log Category was not specified"
CMWT_VALIDATE RmvSet, "Administrative Action option was not selected"

Select Case Ucase(KeySet)
	Case "EVENTS"
		If RmvSet = "-1" Then
			query = "DELETE FROM dbo.EventLog"
			LogDescription = "All [" & KeySet & "] log entries were cleared by " & CMWT_USERNAME()
		Else
			query = "DELETE FROM dbo.EventLog WHERE (EventDateTime < DATEADD(dd, -" & RmvSet & ", GETDATE()) )"
			LogDescription = "[" & KeySet & "] log entries older than " & RmvSet & " were cleared by " & CMWT_USERNAME()
		End If
	Case "TASKS"
		If RmvSet = "-1" Then
			query = "DELETE FROM dbo.Tasks"
			LogDescription = "All [" & KeySet & "] log entries were cleared by " & CMWT_USERNAME()
		Else
			query = "DELETE FROM dbo.Tasks WHERE (DateTimeCreated < DATEADD(dd, -" & RmvSet & ", GETDATE()) )"
			LogDescription = "[" & KeySet & "] log entries older than " & RmvSet & " were cleared by " & CMWT_USERNAME()
		End If
	Case Else:
		query = ""
End Select

On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeOut = 5
conn.Open Application("DSN_CMWT")
If err.Number <> 0 Then
	CMWT_STOP "database connection failure"
End If
conn.Execute query

CMWT_LogEvent conn, "INFO", "CMWT LOG", LogDescription

conn.Close
Set conn = Nothing

TargetURL = "cmwtlog.asp?l=" & KeySet
Caption   = "Clearing CMWT Log Entries"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 1
%>
