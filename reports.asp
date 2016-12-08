<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: reports.asp
' Author....: David M. Stein
' Date......: 12/07/2016
' Purpose...: reports
'****************************************************************
time1 = Timer
PageTitle = "Reports"

CMWT_NewPage "", "", ""
%>
	<!-- #include file="_sm.asp" -->
	<!-- #include file="_banner.asp" -->
	
	<table class="tfx">
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='report0.asp'" title="Custom Device Query">Custom Device Query</td>
			<td class="m111 w250 pad5" onClick="document.location.href='clientpush.asp'" title="Client Push Installations">Client Push Installs</td>
			<td class="m111 w250 pad5" onClick="document.location.href='cmqueries.asp'" title="Queries">Queries</td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 v10 td6" onClick="document.location.href='customreports.asp'" title="Custom Reports">Custom Reports</td>
			<td class="m111 w250 pad5" onClick="document.location.href='report1.asp?fn=OperatingSystem&fv=Windows 10&m=CONTAINS&of=ComputerName,ADSiteName,Model,OperatingSystem'" title="Windows 10 Clients">Windows 10 Clients</td>
			<td class="m111 w250 pad5" onClick="document.location.href='report1.asp?fn=OperatingSystem&fv=Server 2016&m=CONTAINS&of=ComputerName,ADSiteName,Model,OperatingSystem'" title="Windows Server 2016">Windows Server 2016</td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='sqlreports.asp'" title="SQL Reports">SQL Reports</td>
			<td class="m111 w250 v10 td6"></td>
			<td class="m111 w250 v10 td6"></td>
			<td></td>
		</tr>
		<tr class="h50">
			<td class="m111 w250 pad5" onClick="document.location.href='report1.asp?fn=Model&fv=Virtual&m=CONTAINS&of=ComputerName, Model, OperatingSystem&s=ComputerName'" title="Virtual Machines">Virtual Machines</td>
			<td class="m111 w250 v10 td6"></td>
			<td class="m111 w250 v10 td6"></td>
			<td></td>
		</tr>
	</table>
		
	<% CMWT_Footer() %>
	
</body>

</html>