<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepedit2.asp
' lastupdate..... 12/12/2016
' description.... update custom report query
'-----------------------------------------------------------------------------
Response.Expires = -1
RepID    = CMWT_GET("id", "")
RepName  = CMWT_GET("name", "")
RepQuery = CMWT_GET("q", "")
RepComm  = CMWT_GET("comm", "")
RepType  = CMWT_GET("rtype", "1")
CMWT_VALIDATE RepID, "Report Record ID was not provided"
CMWT_VALIDATE RepName, "Report Name was not provided"
CMWT_VALIDATE RepQuery, "Report Query Statement was not provided"
'CMWT_VALIDATE RepComm, "Comment was not provided"

query = "UPDATE dbo.Reports2 " & _
	"SET ReportName='" & RepName & "'," & _
	"ReportType=" & RepType & "," & _
	"Query='" & Replace(RepQuery,"'","''") & "'," & _
	"Comment='" & RepComm & "', " & _
	"CreatedBy='" & CMWT_USERNAME() & "'," & _
	"DateCreated='" & NOW & "' " & _
	"WHERE ReportID=" & RepID
'response.write query
'response.end
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

targetURL = "sqlreports.asp"

Caption = "Updating SQL Report"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 2
'-----------------------------------------------------------------------------
%>
