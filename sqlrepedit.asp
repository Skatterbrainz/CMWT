<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepedit.asp
' lastupdate..... 12/12/2016
' description.... edit a custom SQL report query
'-----------------------------------------------------------------------------
time1 = Timer

ReportID = CMWT_GET("id", "")
CMWT_VALIDATE ReportID, "Report Record ID number was not provided"

PageTitle    = "Modify Report"
PageBackLink = "sqlreports.asp"
PageBackName = "Saved Reports"

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
<form name="form1" id="form1" method="post" action="sqlrepedit2.asp">
<input type="hidden" name="id" id="id" value="<%=ReportID%>" />
<input type="hidden" name="rtype" id="rtype" value="<%=r2%>" />
<table class="tfx">
	<tr>
		<td class="td6a v10 w200 bgBlue">Report Name</td>
		<td class="td6a v10">
			<input type="text" name="name" id="name" class="w500 pad5 v10" title="Enter Report Name" maxlength="50" value="<%=r1%>" />
		</td>
	</tr>
	<tr>
		<td class="td6a v10 w200 bgBlue">Description</td>
		<td class="td6a v10">
			<input type="text" name="comm" id="comm" class="w500 pad5 v10" title="Enter Description" maxlength="255" value="<%=r4%>" />
		</td>
	</tr>
	<tr>
		<td class="td6a v10 w200 bgBlue">SQL Query</td>
		<td class="td6a v10">
			<textarea name="q" id="q" class="w500 h150 pad6 v10" title="Paste SQL Query Statement"><%=r3%></textarea>
			<br/>(2000 character limit. Avoid special characters)
		</td>
	</tr>
</table>
<br/>
<div class="tfx">
	<input type="button" name="b0" id="b0" class="btx w140 h32" value="Cancel" onClick="javascript:history.back(1);" />
	<input type="button" name="b2" id="b2" class="btx w140 h32" value="Clear" onClick="document.location.href='sqlrepedit.asp?id=<%=ReportID%>'" />
	<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Save" />
</div>
</form>

<% CMWT_Footer() %>
	
</body>
</html>