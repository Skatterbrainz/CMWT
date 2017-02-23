<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepdel.asp
' lastupdate..... 02/22/2017
' description.... delete a custom SQL report query
'-----------------------------------------------------------------------------
time1 = Timer

ReportID = CMWT_GET("id", "")
CMWT_VALIDATE ReportID, "Report Record ID number was not provided"

PageTitle    = "Delete Report"
PageBackLink = "sqlreports.asp"
PageBackName = "Saved Reports"

IDColumn     = "ReportID"
TableName    = "Reports2"
TargetURL    = "sqlreports.asp"

CMWT_NewPage "document.form1.name.focus()", "", ""

query = "SELECT TOP 1 " & _
	"ReportID, ReportType, ReportName, Query, CreatedBy, DateCreated, Comment " & _
	"FROM dbo.Reports2 " & _
	"WHERE ReportID = " & ReportID
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
r1 = rs.Fields("ReportName").value
r2 = rs.Fields("ReportType").value
r3 = rs.Fields("Query").value
r4 = rs.Fields("Comment").value
CMWT_DB_CLOSE()

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<form name="form1" id="form1" method="post" action="delrec.asp">
<input type="hidden" name="id" id="id" value="<%=ReportID%>" />
<input type="hidden" name="tn" id="tn" value="<%=TableName%>" />
<input type="hidden" name="pk" id="pk" value="<%=IDColumn%>" />
<input type="hidden" name="t" id="t" value="<%=TargetURL%>" />
<table class="tfx">
	<tr>
		<td class="td6a v10 w200 bgBlue">Report Name</td>
		<td class="td6a v10"><%=r1%></td>
	</tr>
	<tr>
		<td class="td6a v10 w200 bgBlue">Description</td>
		<td class="td6a v10"><%=r4%></td>
	</tr>
	<tr>
		<td class="td6a v10 w200"> </td>
		<td class="td6a v10">
			<p class="cRed">
				<strong>Are you sure you wish to delete this report?</strong>
			</p>
		</td>
	</tr>
</table>
<br/>
<div class="tfx">
	<input type="button" name="b0" id="b0" class="btx w140 h32" value="No" onClick="javascript:history.back(1);" />
	<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Yes" />
</div>
</form>

<% CMWT_Footer() %>
	
</body>
</html>