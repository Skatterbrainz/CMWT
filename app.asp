<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... app.asp
' lastupdate..... 12/09/2016
' description.... computers with given software product installed
'-----------------------------------------------------------------------------
time1 = Timer

pn = CMWT_GET("pn", "")
SortBy  = CMWT_GET("s", "ComputerName")
QueryOn = CMWT_GET("qq", "")

CMWT_VALIDATE pn, "No product name was provided"

PageTitle = "Computers with: " & pn
PageBackLink = "products.asp?ch=" & Left(pn,1)
PageBackName = "Installed Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"dbo.v_R_System.Name0 AS ComputerName,  " & _
	"dbo.v_R_System.ResourceID, " & _
	"dbo.v_R_System.AD_Site_Name0 AS ADSiteName,  " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model,  " & _
	"dbo.v_GS_COMPUTER_SYSTEM.SystemType0 AS SystemType,  " & _
	"dbo.v_GS_OPERATING_SYSTEM.Caption0 AS WindowsType " & _
	"FROM dbo.v_GS_COMPUTER_SYSTEM INNER JOIN " & _
	"dbo.v_R_System ON dbo.v_GS_COMPUTER_SYSTEM.ResourceID = dbo.v_R_System.ResourceID INNER JOIN " & _
	"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID " & _
	"WHERE (dbo.v_GS_COMPUTER_SYSTEM.ResourceID IN " & _
	"(SELECT DISTINCT ResourceID " & _
	"FROM dbo.v_GS_INSTALLED_SOFTWARE_CATEGORIZED " & _
	"WHERE (ARPDisplayName0 = '" & pn & "'))) " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "app.asp?pn=" & pn, ""
CMWT_DB_CLOSE()

Response.Write "<br/><table class=""tfx"">" & _
	"<tr><td class=""td6 v10"">" & _
	"Note: This report queries only by Application Product Name.  In some situations, the same product may appear " & _
	"from multiple Publisher names.</td></tr></table>"

CMWT_SHOW_QUERY()
CMWT_FOOTER() 
Response.Write "</body></html>"
%>