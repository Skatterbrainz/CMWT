<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... admin.asp
' lastupdate..... 12/10/2016
' description.... CMWT administration page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Administration"

If Not CMWT_ADMIN() Then
	Response.Redirect "error.asp?m=Access Denied / Unauthorized User"
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6" colspan="5">
			Insight
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='cmwtlog.asp?l=events'" title="CMWT Event Logs">CMWT Event Logs</td>
		<td class="m111 w250" onClick="document.location.href='cmwtlogs.asp'" title="CMWT Logs">CMWT Logs Summary</td>
		<td class="m111 w250" onClick="document.location.href='cmwtlog.asp?l=tasks'" title="CMWT Task Logs">CMWT Task Logs</td>
		<td class="m111 w250" onClick="document.location.href='notes.asp'" title="CMWT Notes Library">CMWT Notes Library</td>
		<td></td>
	</tr>
	<tr>
		<td class="td6" colspan="5">
			Monitoring
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='activeusers.asp'" title="CMWT Active Users">CMWT Active Users</td>
		<td class="m111 w250" onClick="document.location.href='diag.asp'" title="CMWT Diagnostics">CMWT Diagnostics</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
</table>

<% CMWT_Footer() %>

</body>

</html>