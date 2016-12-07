<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dbdefrag.asp
' lastupdate..... 12/04/2016
' description.... defragment site sql server database index
'-----------------------------------------------------------------------------
confirm = CMWT_GET("x", "")
IF Not CMWT_ADMIN() Then
	CMWT_STOP "Access Denied!"
End If

If confirm = "yes" Then
	caption = "Defragmenting SQL Database indexes"
	query = "EXEC sp_MSforeachtable @command1=""print '?' DBCC DBREINDEX ('?', ' ', 80)"";"
	Set conn = Server.CreateObject("ADODB.Connection")
	On Error Resume Next
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	conn.Execute query
	conn.Close
	Set conn = Nothing
Else
	caption = "No confirmation code provided."
End If

CMWT_PageRedirect "dbstatus.asp", 5
%>
