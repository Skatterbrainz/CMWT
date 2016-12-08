<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... noteadd.asp
' lastupdate..... 11/30/2016
' description.... create custom note attachment
'-----------------------------------------------------------------------------
Response.Expires = -1

PageTitle = "New Note"
SelfLink  = "noteadd.asp"
ItemID    = CMWT_GET("id", "")
ItemType  = CMWT_GET("t", "")
QueryON   = CMWT_GET("qq", "")

CMWT_VALIDATE ItemID, "Item Name or ID was not provided"
CMWT_VALIDATE ItemType, "Item Class or Type was not specified"

time1 = Timer

CMWT_NewPage "document.form1.comm.focus()", "", ""

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<form name="form1" id="form1" method="post" action="noteadd2.asp">
<table class="tfx">
	<tr>
		<td class="td6 v10">
			<textarea name="comm" id="comm" class="w800 h200 pad6"></textarea>
			<br/>(2000 character limit. Avoid special characters)
		</td>
	</tr>
</table>
<br/>
<div class="tfx">
	<input type="hidden" name="id" id="id" value="<%=ItemID%>" />
	<input type="hidden" name="type" id="type" value="<%=ItemType%>" />
	<input type="button" name="b0" id="b0" class="btx w140 h32" value="Cancel" onClick="javascript:history.back(1);" />
	<input type="reset" name="b2" id="b2" class="btx w140 h32" value="Clear" />
	<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Save" />
</div>
</form>

<% CMWT_Footer() %>
	
</body>
</html>