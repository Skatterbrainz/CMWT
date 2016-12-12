<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... assets.asp
' lastupdate..... 12/10/2016
' description.... assets and compliance landing page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle    = "Assets"
PageBackLink = ""
PageBackName = ""
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6 v10" colspan="5">
			Management
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='cmusers.asp'" title="Users">Users</td>
		<td class="m111 w250" onClick="document.location.href='devices.asp'" title="Devices">Devices</td>
		<td class="m111 w250" onClick="document.location.href='mydevice.asp'" title="My Device Info">My Device Info</td>
		<td class="m111 w250" onClick="document.location.href='clients.asp'" title="Client Summary">Client Summary</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='collections.asp?ks=1'" title="User Collections">User Collections</td>
		<td class="m111 w250" onClick="document.location.href='collections.asp?ks=2'" title="Device Collections">Device Collections</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr>
		<td class="td6 v10" colspan="5">
			Inventory
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=2&ch=all'" title="Windows Server Devices">Windows Server Devices</td>
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=4&ch=all'" title="Windows Desktops">Windows Desktops</td>
		<td class="m111 w250" onClick="document.location.href='chassis.asp'" title="Devices by Form Factor">Devices by Form Factor</td>
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=6&ch=all'" title="Physical Clients">Physical Clients</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=3&ch=all'" title="Windows Client Devices">Windows Client Devices</td>
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=5&ch=all'" title="Windows Laptops">Windows Laptops</td>
		<td class="m111 w250" onClick="document.location.href='vmhosts.asp'" title="Virtual Machine Hosts">Virtual Hosts</td>
		<td class="m111 w250" onClick="document.location.href='devices.asp?ks=7&ch=all'" title="Virtual Clients">Virtual Clients</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='mfrs.asp'" title="Devices by Manufacturer">Devices by Manufacturer</td>
		<td class="m111 w250" onClick="document.location.href='models.asp'" title="Devices by Model">Devices by Model</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
</table>

<% CMWT_Footer() %>

</body>

</html>