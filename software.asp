<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... software.asp
' lastupdate..... 12/04/2016
' description.... software information landing page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='packages.asp'" title="Packages">Packages</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='oslist.asp'" title="Operating Systems">Operating Systems</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='products.asp'" title="Installed Software">Installed Software</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='applications.asp'" title="Applications">Applications</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='adkcomps.asp'" title="Boot Image Components">Boot Image Components</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='appvendors.asp'" title="Applications by Vendor">Applications by Vendor</td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='deployments.asp'" title="Deployments">Deployments</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='ie.asp'" title="IE Versions">IE Versions</td>
		<td class="m111 w250 v10 td6"></td>
		<td></td>
	</tr>
	<tr class="h50">
		<td class="m111 w250 v10 td6" onClick="document.location.href='depsummary.asp'" title="Deployment Summary">Deployment Summary</td>
		<td class="m111 w250 v10 td6" onClick="document.location.href='office.asp'" title="Office Versions">Office Versions</td>
		<td class="m111 w250 v10 td6"></td>
		<td></td>
	</tr>
</table>

<% CMWT_FOOTER() %>

</body>

</html>