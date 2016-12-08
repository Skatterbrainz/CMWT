<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... reportsave.asp
' lastupdate..... 12/03/2016
' description.... save a new custom report
'-----------------------------------------------------------------------------
Response.Expires = -1

SearchField  = CMWT_GET("r1","")
SearchValue  = CMWT_GET("r2","")
SearchMode   = CMWT_GET("r3","")
OutputFields = CMWT_GET("r4","")
ReportName   = CMWT_GET("r0","")
Comment      = CMWT_GET("comm","")

CMWT_VALIDATE ReportName, "Report Name was not provided"
CMWT_VALIDATE SearchField, "Search Field was not selected"
CMWT_VALIDATE SearchValue, "Search Value was not specified"
CMWT_VALIDATE SearchMode, "Search Mode was not selected"
CMWT_VALIDATE OutputFields, "Output fields were not selected"

query = "INSERT INTO dbo.Reports " & _
	"(ReportName, SearchField, SearchValue, SearchMode, DisplayColumns, Comment, DateCreated, CreatedBy) " & _
	"VALUES (" & _
	"'" & ReportName & "','" & SearchField & "','" & SearchValue & "','" & SearchMode & _
	"','" & OutputFields & "','" & Comment & "','" & NOW & "','" & CMWT_USERNAME() & "')"

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

targetURL = "reports.asp"

Caption = "Saving new Report"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 1
'-----------------------------------------------------------------------------
%>
