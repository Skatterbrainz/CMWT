<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... software.asp
' lastupdate..... 12/12/2016
' description.... software information landing page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6 v10" colspan="5">
			Software Deployment
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='packages.asp'" title="Packages">Packages</td>
		<td class="m111 w250" onClick="document.location.href='deployments.asp'" title="Deployments">Deployments</td>
		<td class="m111 w250" onClick="document.location.href='updates.asp'" title="Software Updates">Software Updates</td>
		<td class="m111 w250" onClick="document.location.href='depsummary.asp'" title="Deployment Summary">Deployment Summary</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='applications.asp'" title="Applications">Applications</td>
		<td class="m111 w250"></td>
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
		<td class="m111 w250" onClick="document.location.href='products.asp'" title="Installed Software">Installed Software</td>
		<td class="m111 w250" onClick="document.location.href='ie.asp'" title="IE Versions">IE Versions</td>
		<td class="m111 w250" onClick="document.location.href='report1.asp?fn=OperatingSystem&fv=Windows 10&m=CONTAINS&of=ComputerName,ADSiteName,Model,OperatingSystem'" title="Windows 10 Clients">Windows 10 Clients</td>
		<td class="m111 w250" onClick="document.location.href='oslist.asp'" title="Operating Systems">Operating Systems</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='appvendors.asp'" title="Applications by Vendor">Applications by Vendor</td>
		<td class="m111 w250" onClick="document.location.href='office.asp'" title="Office Versions">Office Versions</td>
		<td class="m111 w250" onClick="document.location.href='report1.asp?fn=OperatingSystem&fv=Server 2016&m=CONTAINS&of=ComputerName,ADSiteName,Model,OperatingSystem'" title="Windows Server 2016">Windows Server 2016</td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
	<tr>
		<td class="td6 v10" colspan="5">
			Operating Systems Deployment
		</td>
	</tr>
	<tr class="h50">
		<td class="m111 w250" onClick="document.location.href='bootimages.asp'" title="Boot Image">Boot Images</td>
		<td class="m111 w250" onClick="document.location.href='adkcomps.asp'" title="Boot Image Components">Boot Image Components</td>
		<td class="m111 w250" onClick="document.location.href='tasksequences.asp'" title="Task Sequences">Task Sequences</td>
		<td class="m111 w250"></td>
		<td></td>
	</tr>
</table>

<% CMWT_FOOTER() %>

</body>

</html>