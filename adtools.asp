<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adtools.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: about
'****************************************************************
time1 = Timer
PageTitle = "Active Directory"

CMWT_NewPage "", "", ""
%>
	<!-- #include file="_sm.asp" -->
	<!-- #include file="_banner.asp" -->
	
	<table class="tfx">
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='adcomputers.asp'" title="Computers">Computers</td>
			<td class="m111 w250 pad5" onClick="document.location.href='adattributes.asp'" title="Schema Attributes">Schema Attributes</td>
			<td class="m111 w250 pad5" onClick="document.location.href='addc.asp'" title="Domain Controllers">Domain Controllers</td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='adusers.asp'" title="AD Users">Users</td>
			<td class="m111 w250 pad5" onClick="document.location.href='adsites.asp'" title="AD Sites">AD Sites</td>
			<td class="m111 w250 pad5" onClick="document.location.href='oulist.asp'" title="AD OU Browser">AD OU Browser</td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='adgroups.asp'" title="AD Groups">Groups</td>
			<td class="m111 w250 pad5"></td>
			<td class="m111 w250 pad5"></td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='adprinters.asp'" title="AD Printers">Printers</td>
			<td class="m111 w250 pad5"></td>
			<td class="m111 w250 pad5"></td>
			<td></td>
		</tr>
	</table>
		
	<% CMWT_Footer() %>
	
</body>

</html>