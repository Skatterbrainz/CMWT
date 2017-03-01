<%
'-----------------------------------------------------------------------------
' filename....... _panel3.asp
' lastupdate..... 02/28/2017
' description.... CMWT home page dashboard panel
'-----------------------------------------------------------------------------
%>
<h2>Site Resources</h2>

<table class="t1x">
	<tr class="tr2" onClick="document.location.href='clients.asp'" title="View Records">
		<td class="td5a v10">Forest: Discovered Computers</td>
		<td class="td5a v10 w80 right"><a href="clients.asp"><%=count_computers%></a></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='adusers.asp?x=1'" title="View Records">
		<td class="td5a v10">Forest: Discovered User Accounts</td>
		<td class="td5a v10 w80 right"><a href="adusers.asp?x=1"><%=count_users%></a></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='adgroups.asp'" title="View Records">
		<td class="td5a v10">Forest: Discovered Groups</td>
		<td class="td5a v10 w80 right"><a href="adgroups.asp"><%=count_groups%></a></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='bgroups.asp'" title="View Records">
		<td class="td5a v10">Site: Site Boundary Groups</td>
		<td class="td5a v10 w80 right"><a href="bgroups.asp"><%=count_bgs%></a></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='dpservers.asp'" title="View Records">
		<td class="td5a v10">Site: Distribution Points</td>
		<td class="td5a v10 w80 right"><a href="dpservers.asp"><%=count_dps%></a></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='products.asp'" title="View Records">
		<td class="td5a v10">Site: Inventoried Applications</td>
		<td class="td5a v10 w80 right">
			<%
			If count_apps > 50 Then
				Response.Write "<a href=""products.asp?ch=A"">" & count_apps & "</a>"
			Else
				Response.Write "<a href=""products.asp?ch=ALL"">" & count_apps & "</a>"
			End If
			%>
		</td>
	</tr>
</table>
