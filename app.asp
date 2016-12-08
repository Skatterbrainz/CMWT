<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... app.asp
' lastupdate..... 11/27/2016
' description.... computers with given software product installed
'-----------------------------------------------------------------------------
time1 = Timer

pn = CMWT_GET("pn", "")
SortBy  = CMWT_GET("s", "ComputerName")
QueryOn = CMWT_GET("qq", "")

CMWT_VALIDATE pn, "No product name was provided"

pageTitle = "Computers with: " & pn

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT DISTINCT dbo.v_R_System.Name0 AS ComputerName, dbo.v_R_System.ResourceID, " & _
	"dbo.v_GS_ADD_REMOVE_PROGRAMS.DisplayName0 AS ProductName, " & _
	"dbo.v_GS_ADD_REMOVE_PROGRAMS.Publisher0 AS Publisher, " & _
	"dbo.v_R_System.AD_Site_Name0 AS ADSiteName, " & _
	"dbo.v_GS_OPERATING_SYSTEM.Caption0 AS WindowsType " & _
	"FROM dbo.v_R_System INNER JOIN " & _
	"dbo.v_GS_ADD_REMOVE_PROGRAMS ON dbo.v_R_System.ResourceID = dbo.v_GS_ADD_REMOVE_PROGRAMS.ResourceID INNER JOIN " & _
	"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID " & _
	"WHERE DisplayName0='" & URLDecode(pn) & "' " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "app.asp?pn=" & pn, ""
CMWT_DB_CLOSE()
	
%>
<br/>
<table class="tfx">
	<tr>
		<td class="td6 v10">
			Note: This report queries only by Application Product Name.  In some situations, the same product may appear
			from multiple Publisher names.
		</td>
	</tr>
</table>

<% 
CMWT_SHOW_QUERY()
CMWT_FOOTER() 
%>

</body>
</html>