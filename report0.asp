<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: report0.asp
' Author....: David M. Stein
' Date......: 12/03/2016
' Purpose...: custom device query
'****************************************************************
time1 = Timer
PageTitle = "Custom Device Query"
PageBackLink = "reports.asp"
PageBackName = "Reports"
CMWT_NewPage "", "", ""

fields = Application("CMWT_CUSTOMREPFIELDS")
matchlist = Application("CMWT_CUSTOMREPMODES")

%>
	<!-- #include file="_sm.asp" -->
	<!-- #include file="_banner.asp" -->
	
	<form name="form1" id="form1" method="post" action="report1.asp">
	<table class="tfz">
		<tr>
			<td class="td6 w150 bgGray">Search Field</td>
			<td class="td6">
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
			<td class="td6">
				<input type="text" name="fv" id="fv" class="w400 pad5 v10" maxlength="50" />
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Search Mode</td>
			<td class="td6">
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
			<td class="td6">
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
				<input type="reset" name="b0" id="b0" class="btx w150 v10 h30" label="Clear Form" />
				<input type="submit" name="b1" id="b1" class="btx w150 v10 h30" label="Run!" />
			</td>
		</tr>
	</table>
	</form>

<% CMWT_FOOTER() %>
</body>
</html>