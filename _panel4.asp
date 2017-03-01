<%
'-----------------------------------------------------------------------------
' filename....... _panel4.asp
' lastupdate..... 02/28/2017
' description.... CMWT home page dashboard panel
'-----------------------------------------------------------------------------
%>
<h2>Status</h2>

<table class="t1x">
	<tr class="tr2" onClick="document.location.href='sitestatus.asp'">
		<td class="td5a v10">Site Status Errors</td>
		<td class="td5a v10 w80 right"><%=count_stat1%></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='compstatus.asp'">
		<td class="td5a v10">Component Status Errors</td>
		<td class="td5a v10 w80 right"><%=count_stat2%></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='collections.asp?ks=2&ch=all'">
		<td class="td5a v10">Device Collections</td>
		<td class="td5a v10 w80 right"><%=count_dcolls%></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='collections.asp?ks=1&ch=all'">
		<td class="td5a v10">User Collections</td>
		<td class="td5a v10 w80 right"><%=count_ucolls%></td>
	</tr>
	<tr class="tr2" onClick="document.location.href='tasksequences.asp'">
		<td class="td5a v10">Task Sequences</td>
		<td class="td5a v10 w80 right"><%=count_tseqs%></td>
	</tr>
	<tr class="tr2">
		<td class="td5a v10">.</td>
		<td class="td5a v10 w80 right"> </td>
	</tr>
</table>
