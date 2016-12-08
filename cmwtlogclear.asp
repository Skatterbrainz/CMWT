<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmwtlogclear.asp
' lastupdate..... 12/04/2016
' description.... cmwt log table clear request
'-----------------------------------------------------------------------------
time1 = Timer
KeySet  = CMWT_GET("l", "")

PageTitle    = "Clear CMWT Log"
PageBackLink = "admin.asp"
PageBackName = "Administration"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<form name="form1" id="form1" method="post" action="cmwtlogclear2.asp">
<table class="tfx">
	<tr>
		<td class="td6 v10 bgGray w180">CMWT Log Group</td>
		<td class="td6 v10 bgBlue">
			<%
			If KeySet <> "" Then
				Response.Write "<select name=""x"" id=""x"" size=""1"" class=""pad5 v10 w200"" disabled>" & _
					"<option value="""" selected>" & Ucase(Left(KeySet,1)) & Lcase(Mid(KeySet,2)) & "</option>" & _
					"</select>" & _
					"<input type=""hidden"" name=""l"" id=""l"" value=""" & KeySet & """ />"
			Else
				Response.Write "<select name=""l"" id=""l"" size=""1"" class=""pad5 v10 w200"">"
				For each lc in Split("Events,Tasks", ",")
					If Lcase(lc) = Lcase(KeySet) Then
						Response.Write "<option value=""" & Lcase(lc) & """ selected>" & lc & "</option>"
					Else
						Response.Write "<option value=""" & Lcase(lc) & """>" & lc & "</option>"
					End If
				Next
				Response.Write "</select>"
			End If
			%>
		</td>
	</tr>
	<tr>
		<td class="td6 v10 bgGray w180">Select Clearing Option</td>
		<td class="td6 v10 bgBlue">
			<select name="x1" id="x1" size="1" class="pad5 v10 w200">
				<option value=""></option>
				<option value="1">Clear Entries Older than 1 Day</option>
				<option value="7">Clear Entries Older than 7 Days</option>
				<option value="14">Clear Entries Older than 14 Days</option>
				<option value="30">Clear Entries Older than 30 Days</option>
				<option value="60">Clear Entries Older than 60 Days</option>
				<option value="180">Clear Entries Older than 180 Days</option>
				<option value="-1">Clear All Entries</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td6 v10 bgBlue" colspan="2">
			WARNING: This action is permanent and cannot be undone.
		</td>
	</tr>
</table>
<p class="tfx">
	<input type="submit" name="b1" id="b1" class="btx w140 h30" value="Continue" title="Continue" />
</p>
</form>

</body>
</html>