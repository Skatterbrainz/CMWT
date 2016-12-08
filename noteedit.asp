<!-- #include file="_core.asp" -->
<%
'****************************************************************
' Filename..: noteedit.asp
' Author....: David M. Stein
' Date......: 12/04/2016
' Purpose...: create new custom note
'****************************************************************
Response.Expires = -1
time1 = Timer

PageTitle = "Edit Note"
SelfLink  = "noteedit.asp"
PageBackLink = "notes.asp"
PageBackName = "Notes Library"
RowID     = CMWT_GET("id", "")
QueryON   = CMWT_GET("qq", "")

CMWT_VALIDATE RowID, "Note Record ID was not provided"

query = "SELECT * FROM dbo.Notes WHERE NoteID=" & RowID
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query
ItemType = rs.Fields("AttachClass").value
NoteText = rs.Fields("Comment").value
ItemID   = rs.Fields("AttachedTo").value
CMWT_DB_CLOSE()

CMWT_NewPage "document.form1.comm.focus()", "", ""
'----------------------------------------------------------------
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<form name="form1" id="form1" method="post" action="noteedit2.asp">
<table class="tfx">
	<tr>
		<td class="td6a v10 bgGray w200">Comment</td>
		<td class="td6a v10">
			<textarea name="comm" id="comm" class="w500 h200 v10 pad6"><%=NoteText%></textarea>
		</td>
	</tr>
</table>
<br/>
<div class="tfx">
	<input type="hidden" name="id" id="id" value="<%=RowID%>" />
	<input type="hidden" name="iid" id="iid" value="<%=ItemID%>" />
	<input type="hidden" name="type" id="type" value="<%=ItemType%>" />
	<input type="button" name="b0" id="b0" class="btx w140 h32" value="Cancel" onClick="javascript:history.back(1);" />
	<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Save" />
</div>
</form>
	
<% CMWT_Footer() %>
	
</body>
</html>