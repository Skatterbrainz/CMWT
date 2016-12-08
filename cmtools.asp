<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmtools.asp
' lastupdate..... 11/30/2016
' description.... tools landing page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "CM Tools"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='diag.asp'" title="CMWT Diagnostics">CMWT Diagnostics</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='colltools.asp'" title="Collection Tools">ConfigMgr Collection Tools</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='cmwtlog.asp?l=tasks" title="CMWT Task Logs">CMWT Task Logs</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='activeusers.asp'" title="CMWT Active Users">CMWT Active Users</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='clienttools.asp'" title="Client Tools">ConfigMgr Client Tools</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='cmwtlog.asp?l=events" title="CMWT Event Logs">CMWT Event Logs</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='notes.asp'" title="CMWT Notes Library">CMWT Notes Library</td>
		<td class="m111 v10 td6" onClick="document.location.href='mydevice.asp'" title="My Device Info">My Device Info</td>
		<td class="m111 w250 v10 td6"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='cmwtlogs.asp'" title="CMWT Logs">CMWT Logs Summary</td>
		<td class="m111 w250 v10 td6"></td>
		<td class="m111 w250 v10 td6"></td>
		<td></td>
	</tr>
</table>

<% CMWT_Footer() %>

</body>

</html>