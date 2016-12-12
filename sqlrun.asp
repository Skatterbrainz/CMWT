<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrun.asp
' lastupdate..... 12/08/2016
' description.... execute selected SQL custom report
'-----------------------------------------------------------------------------
time1 = Timer
ReportID = CMWT_GET("id", "")
CMWT_VALIDATE ReportID, "Report ID was not provided"

QueryON = CMWT_GET("qq", "")

query = "SELECT TOP 1 ReportID, ReportType, ReportName, " & _
	"Query, CreatedBy, DateCreated, Comment " & _
	"FROM dbo.Reports2 " & _
	"WHERE ReportID=" & ReportID

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
ReportName  = rs.Fields("ReportName").value
ReportCode  = rs.Fields("Query").value
DateCreated = rs.Fields("DateCreated").value
CreatedBy   = rs.Fields("CreatedBy").value

PageTitle    = "SQL Report"
PageBackLink = "sqlreports.asp"
PageBackName = "SQL Reports"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<table class="tfx">
	<tr>
		<td class="td5 v8">Report Name</td>
		<td class="td5 v8">Created By</td>
		<td class="td5 v8">Date Created</td>
	</tr>
	<tr>
		<td class="td6a v10"><%=ReportName%></td>
		<td class="td6a v10"><%=CreatedBy%></td>
		<td class="td6a v10"><%=DateCreated%></td>
	</tr>
</table>
<%
query = ReportCode
'Response.Write query
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "sqlrun.asp?id=" & ReportID, ""
CMWT_DB_CLOSE()
'Response.Write "<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h30"" value=""New Report"" onClick=""document.location.href='sqlrepadd.asp'"" /><br/>"
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
