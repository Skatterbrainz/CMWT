<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: cmsite.asp
' Author....: David M. Stein
' Date......: 12/10/2016
' Purpose...: cm site hierarchy landing page
'****************************************************************
time1 = Timer
PageTitle = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6 v10" colspan="5">
			Configuration
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='siteconfig.asp'" title="Site Configuration">Site Configuration</td>
		<td class="m111 w250" onClick="document.location.href='cmroles.asp'" title="Security Roles">Security Roles</td>
		<td class="m111 w250" onClick="document.location.href='discoveries.asp'" title="Discovery Methods">Discovery Methods</td>
		<td class="m111 w250" onClick="document.location.href='summarytasks.asp'" title="Summary Tasks">Summary Tasks</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='sitedef.asp'" title="Site Definition">Site Definition</td>
		<td class="m111 w250" onClick="document.location.href='cmscopes.asp'" title="Security Scopes">Security Scopes</td>
		<td class="m111 w250" onClick="document.location.href='cmaccounts.asp'" title="Site Accounts">Site Accounts</td>
		<td class="m111 w250" onClick="document.location.href='cmtasks.asp'" title="Maintenance Tasks">Maintenance Tasks</td>
		<td></td>
	</tr>
	<tr>
		<td class="td6 v10" colspan="5">
			Hierarchy
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='bgroups.asp'" title="Boundary Groups">Boundary Groups</td>
		<td class="m111 w250" onClick="document.location.href='dpservers.asp'" title="DP Servers">Distribution Points</td>
		<td class="m111 w250" onClick="document.location.href='cmqueries.asp'" title="Queries">Queries</td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='boundaries.asp'" title="Site Boundaries">Site Boundaries</td>
		<td class="m111 w250" onClick="document.location.href='dpgroups.asp'" title="DP Server Groups">Distribution Groups</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr>
		<td class="td6 v10" colspan="5">
			Monitoring
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='sitestatus.asp'" title="Site Status">Site Status</td>
		<td class="m111 w250" onClick="document.location.href='clientsummary.asp'" title="Client Summary">Client Summary</td>
		<td class="m111 w250" onClick="document.location.href='sitelogs.asp'" title="Site Log Files">Site Logs</td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='compstatus.asp'" title="Component Status">Component Status</td>
		<td class="m111 w250" onClick="document.location.href='dbstatus.asp'" title="Database Fragmentation">Database Fragmentation</td>
		<td class="m111 w250" onClick="document.location.href='services.asp'" title="Site Server Windows Services Status">Windows Services</td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
</table>
		
<% CMWT_Footer() %>

</body>
</html>