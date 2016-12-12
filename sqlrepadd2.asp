<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepadd2.asp
' lastupdate..... 12/08/2016
' description.... add new custom note
'-----------------------------------------------------------------------------
Response.Expires = -1
RepName  = CMWT_GET("name", "")
RepQuery = CMWT_GET("q", "")
RepComm  = CMWT_GET("comm", "")
RepType  = CMWT_GET("rt", "1")

CMWT_VALIDATE RepName, "Report Name was not provided"
CMWT_VALIDATE RepQuery, "Report Query Statement was not provided"
'CMWT_VALIDATE RepComm, "Comment was not provided"

query = "INSERT INTO dbo.Reports2 " & _
	"(ReportType,ReportName,Query,CreatedBy,DateCreated,Comment) " & _
	"VALUES (" & RepType & ",'" & RepName & "','" & Replace(RepQuery,"'","''") & _
	"','" & CMWT_USERNAME() & "','" & NOW & "','" & RepComm & "')"
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

Caption = "Adding SQL Report"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 2
'-----------------------------------------------------------------------------
%>
