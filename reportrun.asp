<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... reportrun.asp
' lastupdate..... 12/03/2016
' description.... run a custom report
'-----------------------------------------------------------------------------
ReportID = CMWT_GET("id", "")
RunMode = CMWT_GET("rm", "1")

CMWT_VALIDATE ReportID, "Report ID value was not provided"
	
query = "SELECT TOP 1 SearchField, SearchValue, SearchMode, DisplayColumns " & _
	"FROM dbo.Reports " & _
	"WHERE ReportID=" & ReportID

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query

SearchField  = rs.Fields("SearchField").value
SearchValue  = rs.Fields("SearchValue").value
SearchMode   = rs.Fields("SearchMode").value
OutputFields = rs.Fields("DisplayColumns").value

CMWT_DB_CLOSE()

TargetURL = "report1.asp?fn=" & SearchField & "&fv=" & SearchValue & "&m=" & SearchMode & "&of=" & OutputFields & "&rm=" & RunMode

Response.Redirect TargetURL

%>
