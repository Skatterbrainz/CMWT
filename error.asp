<!-- #include file=_core.asp -->
<%
Msg = CMWT_GET("m", "Unknown Error")
time1 = Timer
PageTitle = "Error Report"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="td6a v10 h200 ctr">
			
			<h2>Error / Exception Report</h2>
			
			<p><%=Msg%></p>
			
			<p><a href="javascript:history.back(1);" title="Go Back">Go Back</a></p>
			
		</td>
	</tr>
</table>

</body>
</html>