<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<!-- #include file=_dashboard1.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... default.asp
' lastupdate..... 03/01/2017
' description.... CMWT home page
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle  = Application("CMWT_SubTitle")
SelfLink   = "default.asp"
IsHomePage = True

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<!-- #include file="_panel1.asp" -->
<!-- #include file="_panel2.asp" -->
<!-- #include file="_panel5.asp" -->

<table class="tfx">
	<tr>
		<td class="v10 vtop w600">
			<!-- #include file="_panel3.asp" -->			
		</td>
		<td class="v10 vtop">
			<!-- #include file="_panel4.asp" -->
		</td>
	</tr>
	<tr>
		<td class="v10 vtop w600 td5a">

			<h2><a href="clientsummary.asp" title="Summary Report">Client Installations</a></h2>

			<table class="t1x">
				<tr>
					<td class="td6 v10 w250 bgDarkGray"><a href="clients.asp" title="Computers with Installed Clients">Installed Clients</a></td>
					<td class="td6 v10 bgDarkGray">
						<% CMWT_TABLE_GRAPH2 count_clients, count_computers %>
					</td>
				</tr>
				<tr>
					<td class="td6 v10 w250 bgDarkGray"><a href="clients.asp?c=0" title="Computers Without Installed Clients">Missing Clients</a></td>
					<td class="td6 v10 bgDarkGray">
						<% CMWT_TABLE_GRAPH2 count_null, count_computers %>
					</td>
				</tr>
				<tr>
					<td class="td6 v10 w250 bgDarkGray"><a href="chassis.asp?t=3" title="Desktops">Desktop Computers</a></td>
					<td class="td6 v10 bgDarkGray">
						<% CMWT_TABLE_GRAPH2 count_dt, count_computers %>
					</td>
				</tr>
				<tr>
					<td class="td6 v10 w250 bgDarkGray"><a href="chassis.asp?t=9,10,14" title="Desktops">Laptop Computers</a></td>
					<td class="td6 v10 bgDarkGray">
						<% CMWT_TABLE_GRAPH2 count_lt, count_computers %>
					</td>
				</tr>
			</table>
		</td>
		<td class="v10 vtop td5a">

			<h2><a href="oslist.asp" title="Summary Report">Operating Systems</a></h2>
			
			<table class="t1x">
				<%
				query = "SELECT Caption0 AS OSCaption, " & _
					"COUNT(DISTINCT Name0) AS QTY " & _
					"FROM (" & q_devices & ") AS T1 " & _
					"WHERE Caption0 IS NOT NULL " & _
					"GROUP BY Caption0 " & _
					"ORDER BY QTY DESC"

				CMWT_DB_QUERY Application("DSN_CMDB"), query

				If Not(rs.BOF And rs.EOF) Then
					xrows = rs.RecordCount
					xcols = rs.Fields.Count

					Do Until rs.EOF
						Response.Write "<tr>" & _
							"<td class=""td6 v10 w300 bgDarkGray"">" & _
								"<a href=""os.asp?on=" & rs.Fields("OSCaption").value & _
								""" title=""Computers with " & rs.Fields("OSCaption").value & """>" & _
								rs.Fields("OSCaption").value & "</a></td>" & _
							"<td class=""td6 v10 bgDarkGray"" nowrap>"
						CMWT_TABLE_GRAPH2 rs.Fields("QTY").value, count_computers
						Response.Write "</td></tr>"
						rs.MoveNext
					Loop
				End If
				
				CMWT_DB_CLOSE()

				%>
			</table>
		</td>
	</tr>
</table>

<% CMWT_Footer() %>
	
</body>
</html>