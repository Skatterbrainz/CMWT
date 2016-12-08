<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: copyDirectCollections1.asp
' Author....: David M. Stein
' Date......: 12/07/2016
' Purpose...: copy direct-membership collection assignments from one resource to another
' SQL.......: 
'****************************************************************
Response.Expires = -1
time1 = Timer

Computer1 = CMWT_GET("cn1","")
Computer2 = CMWT_GET("cn2","")

PageTitle = "Copy Computer Collection Assignments"

' PSEUDO CODE: 
' 1. Get list of direct-member collections for Computer1
' 2. Get list of direct-member collections for Computer2 which differ from Computer1
' 3. Present selection list for user to choose from
'    a. Provide check-boxes with "select-all", "clear-all"
' 4. submit ==> go to copyDirectCollections2.asp

CMWT_DB_QUERY Application("DSN_CMDB"), query

'----------------------------------------------------------------
%>

	<table width="100%" border="0" cellpadding="2" cellspacing="2">
		<tr>
			<td style="width:50%;vertical-align:top;">
				<form name="form1" id="form1" method="post" action="">
					Source: <select name="a1" id="a1" size="1" style="padding:5px" onChange="if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }">
						<% CMWT_CM_CLIENTSLIST conn, "copyAssetCollectionsLoader.asp?cn2=" & Computer2 & "&cn1=", Computer1 %>
					</select>
				</form>	
			</td>
			<td style="width:50%;vertical-align:top;">
				<form name="form2" id="form2" method="post" action="">
					Destination: <select name="a2" id="a2" size="1" style="padding:5px" onChange="if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }">
						<% CMWT_CM_CLIENTSLIST conn, "copyAssetCollectionsLoader.asp?cn1=" & Computer1 & "&cn2=", Computer2 %>
					</select>
				</form>
			</td>
		</tr>
		<tr>
			<td style="width:50%;vertical-align:top;">
				<% 
				If q1 <> "" Then 
					If Computer2 <> "" Then
						Response.Write "<form name=""form11"" id=""form11"" method=""post"" action=""copyDirectCollections2.asp"">"
					End If
				%>
				<table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#c0c0c0">
					<tr>
						<td class="v10 bgLightBlue" colspan="2">Copy From...<%=Computer1%></td>
					</tr>
					<tr>
						<td class="v10c bgLightGray" style="width:80px">Coll ID</td>
						<td class="v10 bgLightGray">Collection Name</td>
					</tr>
					<%
					'----------------------------------------------------------------

					If found1 = True Then

						clist1 = ""
						avail = 0

						Do Until rs1.EOF

							cid1 = rs1.Fields("CollectionID").value
							cnm1 = rs1.Fields("Name").value

							If Computer2 <> "" And InStr(Ucase(clist2), Ucase(cid1)) > 0 Then
								cbox = "<img src=""images/box_none.png"" border=""0""/>"
							Else
								cbox = "<input type=""checkbox"" name=""c1"" id=""c1"" value=""" & cid1 & """/>"
								avail = avail + 1
							End If

							Response.Write "<tr>"
							Response.Write "<td class=""v8c bgWhite"">" & cid1 & "</td>"
							Response.Write "<td class=""v8 bgWhite"">" & cbox & " " & cnm1 & "</td>"
							Response.Write "</tr>"

							rs1.MoveNext
						Loop

						rs1.Close
						Set rs1 = Nothing

						Response.Write "<tr>"
						Response.Write "<td class=""v10 bgLightGray"" colspan=""2"">" & xrows1 & " collections found. " & _
							"(" & avail & " available for copying)</td>"
						Response.Write "</tr>"

					Else
						Set rs1 = Nothing
						Response.Write "<tr style=""height:100px"">"
						Response.Write "<td class=""v10c bgWhite"" colspan=""2"">No direct-membership collections found</td>"
						Response.Write "</tr>"
					End If			
					'----------------------------------------------------------------				
					%>
				</table>
				<br/>
				<input type="hidden" name="a11" id="a11" value="<%=Computer1%>"/>
				<input type="hidden" name="a22" id="a22" value="<%=Computer2%>"/>
				<input type="hidden" name="r1" id="r1" value="<%=ResourceID1%>"/>
				<input type="hidden" name="r2" id="r2" value="<%=ResourceID2%>"/>
				<% If Computer2 <> "" Then %>
				<input type="submit" name="b11" id="b11" value="Copy" class="b120"/>
				<% Else %>
				<input type="button" name="b11" id="b11" value="Copy" class="b120x" disabled="true"/>
				&nbsp;<span style="color:#c0c0c0">Select Asset 2 to finish copying assignments...</span>
				<% End If %>
				<%=m1%>
			</form>
			<% End If %>
			</td>
			<td style="width:50%;vertical-align:top;">
				<% If q2 <> "" Then %>
				<table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#c0c0c0">
					<tr>
						<td class="v10 bgLightBlue" colspan="2">Copy To...<%=Computer2%></td>
					</tr>
					<tr>
						<td class="v10c bgLightGray" style="width:80px">Coll ID</td>
						<td class="v10 bgLightGray">Collection Name</td>
					</tr>
					<%
					'----------------------------------------------------------------

					If found2 = True Then

						Do Until rs2.EOF

							cid2 = rs2.Fields("CollectionID").value
							cnm2 = rs2.Fields("Name").value

							Response.Write "<tr>"
							Response.Write "<td class=""v8c bgWhite"">" & cid2 & "</td>"
							Response.Write "<td class=""v8 bgWhite"">" & cnm2 & "</td>"
							Response.Write "</tr>"

							rs2.MoveNext
						Loop

						rs2.Close
						Set rs2 = Nothing

						Response.Write "<tr>"
						Response.Write "<td class=""v10 bgLightGray"" colspan=""2"">" & xrows2 & " collections found</td>"
						Response.Write "</tr>"

					Else
						Response.Write "<tr style=""height:100px"">"
						Response.Write "<td class=""v10c bgWhite"" colspan=""2"">No direct-membership collections found</td>"
						Response.Write "</tr>"
					End If			
					'----------------------------------------------------------------				
					%>
				</table>
				<% End If %>
				<%=m2%>
			</td>
		</tr>
	</table>

</body>
</html>