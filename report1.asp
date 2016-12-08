<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... report1.asp
' lastupdate..... 12/03/2016
' description.... custom device query output
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Device Reports"
PageBackLink = "report0.asp"
PageBackName = "Custom Device Query"

SortBy       = CMWT_GET("s", "ComputerName")
QueryOn      = CMWT_GET("qq", "")
OutputFields = CMWT_GET("of", "")
SearchField  = CMWT_GET("fn", "")
SearchValue  = CMWT_GET("fv", "")
SearchMode   = CMWT_GET("m", "EQUALS")
RunMode      = CMWT_GET("rm", "1")

CMWT_NewPage "", "", ""

%>
<!-- #include file="_sm.asp" -->
<table class="tfx">
	<tr>
		<td class="v10 pad6 bgDarkGray">
			<p>QUERY: Where [<%=SearchField%>] [<%=SearchMode%>] [<%=SearchValue%>]</p>
			<p>SORT BY: [<%=SortBy%>]</p>
		</td>
	</tr>
</table>
<%

baseQuery = "SELECT " & _
	"dbo.v_R_System.ResourceID, " & _
	"dbo.v_R_System.Name0 AS ComputerName, " & _
	"dbo.v_R_System.AD_Site_Name0 AS ADSiteName, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.SystemType0 AS SystemType, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Domain0 AS Domain, " & _
	"dbo.v_GS_OPERATING_SYSTEM.Caption0 AS OperatingSystem, " & _
	"dbo.v_GS_X86_PC_MEMORY.TotalPhysicalMemory0 AS PhysicalMemory, " & _
	"dbo.v_GS_SYSTEM_ENCLOSURE.SerialNumber0 AS SerialNumber, " & _
	"dbo.vMacAddresses.MacAddresses AS MAC, " & _
	"dbo.v_RA_System_SystemOUName.System_OU_Name0 AS OUName, " & _
	"dbo.v_GS_WORKSTATION_STATUS.LastHWScan, " & _
	"dbo.v_GS_SYSTEM.SystemRole0 AS Role, " & _
	"dbo.v_GS_PROCESSOR.Name0 AS Processor, " & _
	"dbo.v_GS_LastSoftwareScan.LastScanDate AS LastSWScan, " & _
	"dbo.v_R_System.Client_Version0 AS ClientVersion " & _
	"FROM dbo.v_R_System INNER JOIN " & _
	"dbo.v_GS_COMPUTER_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_COMPUTER_SYSTEM.ResourceID " & _
	"INNER JOIN " & _
	"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID " & _
	"INNER JOIN " & _
	"dbo.v_GS_X86_PC_MEMORY ON dbo.v_R_System.ResourceID = dbo.v_GS_X86_PC_MEMORY.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_GS_LastSoftwareScan ON dbo.v_R_System.ResourceID = dbo.v_GS_LastSoftwareScan.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_GS_PROCESSOR ON dbo.v_R_System.ResourceID = dbo.v_GS_PROCESSOR.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_GS_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_SYSTEM.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_GS_WORKSTATION_STATUS ON dbo.v_R_System.ResourceID = dbo.v_GS_WORKSTATION_STATUS.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_RA_System_SystemOUName ON dbo.v_R_System.ResourceID = dbo.v_RA_System_SystemOUName.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.v_GS_SYSTEM_ENCLOSURE ON dbo.v_R_System.ResourceID = dbo.v_GS_SYSTEM_ENCLOSURE.ResourceID " & _
	"LEFT OUTER JOIN " & _
	"dbo.vMacAddresses ON dbo.v_R_System.ResourceID = dbo.vMacAddresses.ItemKey"

If OutputFields <> "" Then
	query = "SELECT DISTINCT " & OutputFields & " FROM (" & baseQuery & ") AS T1 "
Else
	query = "SELECT DISTINCT * FROM (" & baseQuery & ") AS T1 "
End If

If SearchField <> "" And SearchValue <> "" Then
	Select Case Ucase(SearchMode)
		Case "EQUALS","EXACT"
			query = query & " WHERE (T1." & SearchField & " = '" & SearchValue & "')"
		Case "GREATER","GREATERTHAN"
			query = query & " WHERE (T1." & SearchField & " > '" & SearchValue & "')"
		Case "EQUALORGREATER","GREATEROREQUAL"
			query = query & " WHERE (T1." & SearchField & " >= '" & SearchValue & "')"
		Case "LIKE","CONTAINS"
			query = query & " WHERE (T1." & SearchField & " LIKE '%" & SearchValue & "%')"
		Case "BEGINS","BEGINSWITH","STARTS","STARTSWITH"
			query = query & " WHERE (T1." & SearchField & " LIKE '" & SearchValue & "%')"
		Case "ENDS","ENDSWITH"
			query = query & " WHERE (T1." & SearchField & " LIKE '%" & SearchValue & "')"
		Case "NOTEQUALS","ISNOT"
			query = query & " WHERE (T1." & SearchField & " <> '" & SearchValue & "')"
		Case "NOTLIKE","NOTCONTAINS"
			query = query & " WHERE (T1." & SearchField & " NOT LIKE '%" & SearchValue & "%')"
	End Select
End If

query = query & " ORDER BY T1." & SortBy

'response.write "<p>field: " & SearchField & "</p>"
'response.write "<p>value: " & SearchValue & "</p>"
'response.write "<p>match: " & SearchMode & "</p>"
'response.write "<p>output: " & OutputFields & "</p>"
'response.write CMWT_PrettySQL(query)
'response.end

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "report1.asp?fn=" & SearchField & "&fv=" & SearchValue & "&m=" & SearchMode & "&of=" & OutputFields, "COMPUTERNAME^device.asp?cn="
CMWT_DB_CLOSE()
if RunMode <> "0" then
%>
<form name="form1" id="form1" method="post" action="reportsave.asp">
<h2>Save Report</h2>
<table class="tf800" style="margin-left:0">
	<tr>
		<td class="td6 v10">Report Name</td>
		<td class="td6 v10">Comment</td>
		<td></td>
	</tr>
	<tr>
		<td class="td6 v10">
			<input type="hidden" name="r1" id="r1" value="<%=SearchField%>" />
			<input type="hidden" name="r2" id="r2" value="<%=SearchValue%>" />
			<input type="hidden" name="r3" id="r3" value="<%=SearchMode%>" />
			<input type="hidden" name="r4" id="r4" value="<%=OutputFields%>" />
			<input type="text" name="r0" id="r0" size="30" maxlength="50" class="pad5 v10" title="Report Name" /> 
		</td>
		<td class="td6 v10">
			<input type="text" name="comm" id="comm" size="50" maxlength="255" class="pad5 v10" title="Comment" />
		</td>
		<td>
			<input type="submit" name="b1" id="b1" class="btx w140 h30" value="Save" title="Save" />
		</td>
	</tr>
</table>
</form>
<%
end if
CMWT_SHOW_Query()
CMWT_Footer()
%>

</body>
</html>