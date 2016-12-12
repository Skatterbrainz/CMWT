<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: report0.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: custom device query
'****************************************************************
time1 = Timer

ReportID = CMWT_GET("id", "")
CMWT_VALIDATE ReportID, "Report ID value was not provided"

PageTitle = "Edit Device Query"
PageBackLink = "customreports.asp"
PageBackName = "Custom Reports"
CMWT_NewPage "", "", ""

query = "SELECT TOP 1 ReportName, SearchField, SearchValue, SearchMode, DisplayColumns, Comment " & _
	"FROM dbo.Reports " & _
	"WHERE ReportID=" & ReportID

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query

ReportName   = rs.Fields("ReportName").value
SearchField  = rs.Fields("SearchField").value
SearchValue  = rs.Fields("SearchValue").value
SearchMode   = rs.Fields("SearchMode").value
OutputFields = rs.Fields("DisplayColumns").value
Comment      = rs.Fields("Comment").value

CMWT_DB_CLOSE()

fields = Application("CMWT_CUSTOMREPFIELDS")
matchlist = Application("CMWT_CUSTOMREPMODES")

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<br/>
<form name="form1" id="form1" method="post" action="reportedit2.asp">
	<table class="tfx">
		<tr>
			<td class="td6 w150 bgGray">Report Name</td>
			<td class="td6 bgMedGray">
				<input type="text" name="rn" id="rn" class="w400 pad5 v10" maxlength="50" value="<%=ReportName%>" title="Name of this Report" />
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Comment</td>
			<td class="td6 bgMedGray">
				<input type="text" name="comm" id="comm" class="w400 pad5 v10" maxlength="255" value="<%=Comment%>" title="Comment or Description" />
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Search Field</td>
			<td class="td6 bgMedGray">
				<select name="r1" id="r1" size="1" class="w180 pad5" title="Select Field to base Query upon">
					<option value=""></option>
					<%
					For each fn in Split(fields,",")
						if Lcase(fn) = Lcase(SearchField) then
							Response.Write "<option value=""" & fn & """ selected>" & fn & "</option>"
						else
							Response.Write "<option value=""" & fn & """>" & fn & "</option>"
						end if
					Next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Search Value</td>
			<td class="td6 bgMedGray">
				<input type="text" name="r2" id="r2" class="w400 pad5 v10" maxlength="50" value="<%=SearchValue%>" title="Enter a Search Value" />
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Search Mode</td>
			<td class="td6 bgMedGray">
				<select name="r3" id="r3" size="1" class="w180 pad5" title="Select search condition mode">
					<%
					For each fn in Split(matchlist,",")
						if Lcase(fn) = Lcase(SearchMode) then
							Response.Write "<option value=""" & fn & """ selected>" & fn & "</option>"
						else
							Response.Write "<option value=""" & fn & """>" & fn & "</option>"
						end if
					Next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td class="td6 w150 bgGray">Output Fields</td>
			<td class="td6 bgMedGray">
				<select name="r4" id="r4" size="8" class="w180 pad5" multiple="true" title="Select Fields to show in output">
					<%
					For each fn in Split(fields,",")
						if InStr(OutputFields, fn) > 0 then
							Response.Write "<option value=""" & fn & """ selected>" & fn & "</option>"
						else
							Response.Write "<option value=""" & fn & """>" & fn & "</option>"
						end if
					Next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<input type="hidden" name="id" id="id" value="<%=ReportID%>" />
				<input type="button" name="b0" id="b0" class="btx w140 v10 h30" value="Cancel" onClick="document.location.href='customreports.asp'" title="Cancel" />
				<input type="submit" name="b1" id="b1" class="btx w140 v10 h30" value="Save" title="Save" />
			</td>
		</tr>
	</table>
</form>

<% CMWT_FOOTER() %>
</body>
</html>