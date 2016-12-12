<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: report0.asp
' Author....: David M. Stein
' Date......: 12/09/2016
' Purpose...: custom device query
'****************************************************************
time1 = Timer
PageTitle    = "Report Builder"
PageBackLink = "reports.asp"
PageBackName = "Reports"
CMWT_NewPage "", "", ""

fields = Application("CMWT_CUSTOMREPFIELDS")
matchlist = Application("CMWT_CUSTOMREPMODES")

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="top w600">
			<h2>Query Builder</h2>
			<form name="form1" id="form1" method="post" action="report1.asp">
				<table class="tfz">
					<tr>
						<td class="td6 w150 bgGray">Search Field</td>
						<td class="td6 bgMedGray">
							<select name="fn" id="fn" size="1" class="w180 pad5">
								<option value=""></option>
								<%
								For each fn in Split(fields,",")
									Response.Write "<option value=""" & fn & """>" & fn & "</option>"
								Next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td6 w150 bgGray">Search Value</td>
						<td class="td6 bgMedGray">
							<input type="text" name="fv" id="fv" class="w400 pad5 v10" maxlength="50" />
						</td>
					</tr>
					<tr>
						<td class="td6 w150 bgGray">Search Mode</td>
						<td class="td6 bgMedGray">
							<select name="m" id="m" size="1" class="w180 pad5">
								<%
								For each fn in Split(matchlist,",")
									Response.Write "<option value=""" & fn & """>" & fn & "</option>"
								Next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td6 w150 bgGray">Output Fields</td>
						<td class="td6 bgMedGray">
							<select name="of" id="of" size="8" class="w180 pad5" multiple="true">
								<%
								For each fn in Split(fields,",")
									Response.Write "<option value=""" & fn & """>" & fn & "</option>"
								Next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<input type="reset" name="b0" id="b0" class="btx w150 v10 h30" value="Clear" />
							<input type="submit" name="b1" id="b1" class="btx w150 v10 h30" value="Execute!" />
						</td>
					</tr>
				</table>
			</form>
		</td>
		<td class="top">
			<h2>Direct SQL Query</h2>
			<form name="form1" id="form1" method="post" action="sqlrepadd2.asp">
				<table class="tfz">
					<tr>
						<td class="td6 w200 bgGray">Report Name</td>
						<td class="td6 bgMedGray">
							<input type="text" name="name" id="name" class="w400 pad5 v10" title="Enter Report Name" maxlength="50" />
						</td>
					</tr>
					<tr>
						<td class="td6 w200 bgGray">Description</td>
						<td class="td6 bgMedGray">
							<input type="text" name="comm" id="comm" class="w400 pad5 v10" title="Enter Description" maxlength="255" />
						</td>
					</tr>
					<tr>
						<td class="td6 bgMedGray" colspan="2">
							<input type="text" name="misc" id="misc" class="w500 pad5 v10" value="Note: 2000 character limit on SQL expression. Avoid special characters" disabled />
						</td>
					</tr>
					<tr>
						<td class="td6 w200 bgGray">SQL Query</td>
						<td class="td6 bgMedGray">
							<textarea name="q" id="q" class="w400 pad6 v10" title="Paste SQL Query Statement" rows="8"></textarea>
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<input type="reset" name="b2" id="b2" class="btx w140 h32" value="Clear" />
							<input type="submit" name="b1" id="b1" class="btx w140 h32" value="Save" />
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>


<% CMWT_FOOTER() %>
</body>
</html>