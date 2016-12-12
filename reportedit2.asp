<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... reportedit2.asp
' lastupdate..... 12/03/2016
' description.... update custom report
'-----------------------------------------------------------------------------
Response.Expires = -1

ReportID  = CMWT_GET("id", "")
SearchField  = CMWT_GET("r1", "")
SearchValue  = CMWT_GET("r2", "")
SearchMode   = CMWT_GET("r3", "")
OutputFields = CMWT_GET("r4", "")
ReportName   = CMWT_GET("rn", "")
Comment      = CMWT_GET("comm", "")

CMWT_VALIDATE ReportID, "Report ID was not provided"
CMWT_VALIDATE ReportName, "Report Name was not provided"
CMWT_VALIDATE SearchField, "Search Field was not selected"
CMWT_VALIDATE SearchValue, "Search Value was not specified"
CMWT_VALIDATE SearchMode, "Search Mode was not selected"
CMWT_VALIDATE OutputFields, "Output fields were not selected"

query = "UPDATE dbo.Reports " & _
	"SET " & _
	"ReportName='" & ReportName & "', " & _
	"SearchField='" & SearchField & "', " & _
	"SearchValue='" & SearchValue & "', " & _
	"SearchMode='" & SearchMode & "', " & _
	"DisplayColumns='" & OutputFields & "', " & _
	"Comment='" & Comment & "', " & _
	"DateCreated='" & NOW & "', " & _
	"CreatedBy='" & CMWT_USERNAME() & "' " & _
	"WHERE ReportID=" & ReportID

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

targetURL = "customreports.asp"

Caption = "Updaing Note Record"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 1
'-----------------------------------------------------------------------------
%>
