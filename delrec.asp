<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... delrec.asp
' lastupdate..... 11/30/2016
' description.... delete database record (generic handler)
'-----------------------------------------------------------------------------
RecordID  = CMWT_GET("id", "")
IDColumn  = CMWT_GET("pk", "")
TableName = CMWT_GET("tn", "")
TargetURL = CMWT_GET("t", "")

CMWT_VALIDATE RecordID, "Table Row ID was not provided"
CMWT_VALIDATE TableName, "Database Table name was not specified"
CMWT_VALIDATE TargetURL, "Target landing document was not specified"

TargetURL = Replace(Replace(TargetURL, "|", "?"), "^", "&")

query = "DELETE FROM dbo." & TableName & " WHERE " & IDColumn & "=" & RecordID

On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeOut = 5
conn.Open Application("DSN_CMWT")
If err.Number <> 0 Then
	CMWT_STOP "database connection failure"
End If
conn.Execute query
conn.Close
Set conn = Nothing

Caption = "Deleting Record"
PageTitle = "ConfigMgr Web Tools"

CMWT_PageRedirect TargetURL, 1

'----------------------------------------------------------------
%>