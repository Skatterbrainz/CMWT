<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepexp.asp
' lastupdate..... 12/10/2016
' description.... render exported sql queries
'-----------------------------------------------------------------------------
query = "SELECT DISTINCT " & _
	"ReportType, ReportName, Query, CreatedBy, DateCreated, Comment " & _
	"FROM dbo.Reports2 " & _
	"ORDER BY ReportName"

Dim conn, cmd, rs
CMWT_DB_OPEN Application("DSN_CMWT")
Set cmd  = Server.CreateObject("ADODB.Command")
Set rs   = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockReadOnly
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = query
rs.Open cmd

Response.Write "INSERT INTO dbo.Reports2<br/>" & _
	"([ReportType],[ReportName],[Query],[CreatedBy],[DateCreated],[Comment])<br/>" & _
	"VALUES <br/>"

rows = rs.RecordCount
crow = 1
Do Until rs.EOF
	Response.Write "(" & rs.Fields("ReportType").value & ", " & _
		"'" & rs.Fields("ReportName").value & "', " & _
		"'" & rs.Fields("Query").value & "', " & _
		"'" & rs.Fields("CreatedBy").value & "', " & _
		"'" & rs.Fields("DateCreated").value & "', " & _
		"'" & rs.Fields("Comment").value & "')"
	If crow < rows Then
		Response.Write ",<br/>"
	Else
		Response.Write "<br/>"
	End If
	crow = crow + 1
	rs.MoveNext
Loop

CMWT_DB_CLOSE()
%>