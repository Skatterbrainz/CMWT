<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<!-- #include file=_m_gen_summary.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... default.asp
' lastupdate..... 03/20/2016
' description.... home page
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle  = Application("CMWT_SubTitle")
SelfLink   = "default.asp"
IsHomePage = True

CMWT_NewPage "", "", ""
'CMWT_PageHeading PageTitle, "" 

%>
	
	<table class="tfx">
		<tr>
			<td>
				<br/>
				<%
				q = "SELECT TOP 1 SiteCode,SiteName,Version," & _
					"ServerName,InstallDir FROM dbo.v_Site"
				CMWT_DB_QUERY Application("DSN_CMDB"), q
				Response.Write "<table class=""t1x""><tr>"
				For i = 0 to rs.Fields.Count - 1
					Response.Write "<td class=""td6a v10 bgBlue"">" & rs.Fields(i).Name & "</td>"
				Next
				Response.Write "<td class=""td6a v10 bgBlue"">Branch Name</td>"
				Response.Write "</tr>"
				Do Until rs.EOF
					Response.Write "<tr>"
					For i = 0 to rs.Fields.Count - 1
						Response.Write "<td class=""td6a v10"">" & rs.Fields(i).Value & "</td>"
					Next
					Response.Write "<td class=""td6a v10"">" & CMWT_CM_BuildName(rs.Fields("Version").value) & "</td>"
					Response.Write "</tr>"
					rs.MoveNext
				Loop
				'CMWT_DB_TABLEGRID rs, "", "", ""
				CMWT_DB_CLOSE()
				Response.Write "</table>"
				%>
			</td>
		</tr>
	</table>
	
	<table class="tfx">
		<tr>
			<td class="v10 vtop w500">
				
				<h2>Site Resources</h2>
				
				<table class="t1x">
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx4=clients.asp'" title="View Records">
						<td class="td5a v10 bgBlue">Forest: Discovered Computers</td>
						<td class="td5a v10 bgDarkGray w80 right"><a href="./?sbx4=clients.asp"><%=count_computers%></a></td>
					</tr>
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx4=adusers.asp&x=1'" title="View Records">
						<td class="td5a v10 bgBlue">Forest: Discovered User Accounts</td>
						<td class="td5a v10 bgDarkGray w80 right"><a href="./?sbx4=adusers.asp?x=1"><%=count_users%></a></td>
					</tr>
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx4=adgroups.asp'" title="View Records">
						<td class="td5a v10 bgBlue">Forest: Discovered Groups</td>
						<td class="td5a v10 bgDarkGray w80 right"><a href="./?sbx4=adgroups.asp"><%=count_groups%></a></td>
					</tr>
				</table>
				
			</td>
			<td class="v10 vtop">
				
				<h2><a href="sitestatus.asp" title="Site Status">Status</a></h2>
				
				<table class="t1x">
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx1=1&sbx2=1&sbx4=bgroups.asp'" title="View Records">
						<td class="td5a v10 bgBlue">Site: Site Boundary Groups</td>
						<td class="td5a v10 bgDarkGray w80 right"><a href="./?sbx1=1&sbx2=1&sbx4=bgroups.asp"><%=count_bgs%></a></td>
					</tr>
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx4=apps.asp'" title="View Records">
						<td class="td5a v10 bgBlue">Site: Inventoried Applications</td>
						<td class="td5a v10 bgDarkGray w80 right">
							<%
							If count_apps > 50 Then
								Response.Write "<a href=""apps.asp?ch=A"">" & count_apps & "</a>"
							Else
								Response.Write "<a href=""apps.asp?ch=ALL"">" & count_apps & "</a>"
							End If
							%>
						</td>
					</tr>
					<tr class="tr1 ptr" onClick="document.location.href='./?sbx1=1&sbx2=1&sbx4=dpservers.asp'" title="View Records">
						<td class="td5a v10 bgBlue w150">Site: Distribution Points</td>
						<td class="td5a v10 bgDarkGray w80 right"><a href="./?sbx1=1&sbx2=1&sbx4=dpservers.asp"><%=count_dps%></a></td>
					</tr>
				</table>
				
			</td>
		</tr>
		<tr>
			<td class="v10 vtop w500 td5a">

				<h2><a href="clientsummary.asp" title="Summary Report">Client Installations</a></h2>

				<table class="t1x">
					<tr>
						<td class="td6 v10 w250 bgDarkGray"><a href="./?sbx4=clients.asp" title="Computers with Installed Clients">Installed Clients</a></td>
						<td class="td6 v10 bgDarkGray">
							<% CMWT_TABLE_GRAPH2 count_clients, count_computers %>
						</td>
					</tr>
					<tr>
						<td class="td6 v10 w250 bgDarkGray"><a href="./?sbx4=clients.asp&c=0" title="Computers Without Installed Clients">Missing Clients</a></td>
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
								"<td class=""td6 v10 w300"">" & _
									"<a href=""os.asp?on=" & rs.Fields("OSCaption").value & _
									""" title=""Computers with " & rs.Fields("OSCaption").value & """>" & _
									rs.Fields("OSCaption").value & "</a></td>" & _
								"<td class=""td6 v10"" nowrap>"
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