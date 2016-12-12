<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sqlrepedit.asp
' lastupdate..... 12/09/2016
' description.... edit a custom SQL report query
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer

PageTitle    = "Modify Report"
PageBackLink = "sqlreports.asp"
PageBackName = "Saved Reports"

CMWT_NewPage "document.form1.name.focus()", "", ""

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<form name="form1" id="form1" method="post" action="sqlrepadd2.asp">
<table class="tfx">
	<tr>
		<td class="td6a v10 w200 bgBlue">Report Name</td>
		<td class="td6a v10">
			<input type="text" name="name" id="name" class="w500 pad5 v10" title="Enter Report Name" maxlength="50" />
		</td>
	</tr>
	<tr>
		<td class="td6a v10 w200 bgBlue">Description</td>
		<td class="td6a v10">
			<input type="text" name="comm" id="comm" class="w500 pad5 v10" title="Enter Description" maxlength="255" />
		</td>
	</tr>
	<tr>
		<td class="td6a v10 w200 bgBlue">SQL Query</td>
		<td class="td6a v10">
			<textarea name="q" id="q" class="w500 h150 pad6 v10" title="Paste SQL Query Statement"></textarea>
			<br/>(2000 character limit. Avoid special characters)
		</td>
	</tr>
</table>
<br/>
<div class="tfx">
	<input type="button" name="b0" id="b0" class="btx w140 h32" value="Cancel" onClick="javascript:history.back(1);" />
	<input type="button" name="b2" id="b2" class="btx w140 h32" value="Clear" onClick="document.location.href='sqlrepadd.asp'" />
	<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Save" />
</div>
</form>

<% CMWT_Footer() %>
	
</body>
</html>