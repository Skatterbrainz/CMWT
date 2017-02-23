<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: confirm.asp
' Author....: David M. Stein
' Date......: 02/22/2017
' Purpose...: generic deletion confirmation prompt
'****************************************************************
Response.Expires = -1

PageTitle = "Confirm Request"
TableName = CMWT_GET("tn", "")
IDColumn  = CMWT_GET("pk", "")
RecordID  = CMWT_GET("id", "")
TargetURL = CMWT_GET("t", "")
QueryON   = CMWT_GET("qq", "")

CMWT_VALIDATE RecordID, "Table Row ID was not provided"
CMWT_VALIDATE IDColumn, "Table Column Identifier was not specified"
CMWT_VALIDATE TableName, "Database Table name was not specified"

time1 = Timer

CMWT_NewPage "document.form1.rn.focus()", "", ""
CMWT_PageHeading PageTitle, ""
'----------------------------------------------------------------
%>

<form name="form1" id="form1" method="post" action="delrec.asp">
<input type="hidden" name="tn" id="tn" value="<%=TableName%>" />
<input type="hidden" name="id" id="id" value="<%=RecordID%>" />
<input type="hidden" name="pk" id="pk" value="<%=IDColumn%>" />
<input type="hidden" name="t" id="t" value="<%=TargetURL%>" />
<table class="t1000x">
	<tr class="h300">
		<td class="td6 v10 ctr">
			<br/>
			<strong>Are you sure you wish to delete this record?</strong>
			<br/><br/>
			<input type="button" name="b0" id="b0" class="btx w140 h32" value="No" onClick="javascript:history.back(1);" />
			<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Yes" />

		</td>
	</tr>
</table>
</form>

<% CMWT_Footer() %>
	
</body>
</html>