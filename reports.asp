<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: reports.asp
' Author....: David M. Stein
' Date......: 01/02/2017
' Purpose...: reports landing page
'****************************************************************
time1 = Timer
PageTitle = "Reports"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6" colspan="5">
			x
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='report0.asp'" title="Report Builder">Report Builder</td>
		<td class="m111 w250" onClick="document.location.href='clientpush.asp'" title="Client Push Installations">Client Push Installs</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='customreports.asp'" title="Custom Reports">Custom Reports</td>
		<td class="m111 w250" onClick="document.location.href='report1.asp?fn=Model&fv=Virtual&m=CONTAINS&of=ComputerName, Model, OperatingSystem&s=ComputerName'" title="Virtual Machines">Virtual Machines</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='sqlreports.asp'" title="Saved SQL Reports">Saved SQL Reports</td>
		<td class="m111 w250" onClick="document.location.href='logins.asp'" title="Device Logins">Device Logins</td>
		<td class="m111 w250"></td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
</table>
	
<% CMWT_Footer() %>
	
</body>

</html>