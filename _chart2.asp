<%
'-----------------------------------------------------------------------------
' filename....... _chart2.asp
' lastupdate..... 02/26/2016
' description.... graphic chart module 2
'-----------------------------------------------------------------------------

Function CMWT_CM_CLIENTCOUNT ()
	Dim query, conn, cmd, rs, result
	query = "SELECT COUNT(*) AS Computers FROM (SELECT DISTINCT ResourceID FROM dbo.v_R_System) AS T1"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = rs.Fields("Computers").value
	Else
		result = 0
	End If
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CM_CLIENTCOUNT = result
End Function

Function CMWT_CM_APPCOUNT ()
	Dim query, conn, cmd, rs, result
	query = "SELECT COUNT(*) AS Apps FROM (SELECT DISTINCT DisplayName0 FROM dbo.v_GS_ADD_REMOVE_PROGRAMS) AS T1"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = rs.Fields("Apps").value
	Else
		result = 0
	End If
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CM_APPCOUNT = result
End Function

Function CMWT_CHART_DATA (query)
	Dim conn, cmd, rs, xx, f1, f2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		xx = ""
		Do Until rs.EOF
			f1 = rs.Fields("ItemName").value
			f2 = rs.Fields("QTY").value
			If xx <> "" Then
				xx = xx & "|" & f2 & "=" & f1
			Else
				xx = f2 & "=" & f1
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CHART_DATA = Split(xx, "|")
End Function

%>