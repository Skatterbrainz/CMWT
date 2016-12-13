<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... chassistype.asp
' lastupdate..... 12/12/2016
' description.... list computers by specific chassis type number
'-----------------------------------------------------------------------------
time1 = Timer
ctnum   = CMWT_GET("ct", "")
SortBy  = CMWT_GET("s", "ComputerName")

CMWT_VALIDATE ctnum, "Chassis Type number was not provided"

PageTitle    = "Chassis Type: " & ctnum
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT dbo.v_GS_COMPUTER_SYSTEM.ResourceID, dbo.v_GS_COMPUTER_SYSTEM.Name0 AS ComputerName, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Domain0 AS Domain, dbo.v_GS_COMPUTER_SYSTEM.Manufacturer0 AS Manufacturer, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model, dbo.v_GS_COMPUTER_SYSTEM.UserName0 AS UserName, " & _
	"dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 AS Chassis " & _
	"FROM dbo.v_GS_SYSTEM_ENCLOSURE INNER JOIN " & _
	"dbo.v_GS_COMPUTER_SYSTEM ON dbo.v_GS_SYSTEM_ENCLOSURE.ResourceID = dbo.v_GS_COMPUTER_SYSTEM.ResourceID " & _
	"WHERE (dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 = " & ctnum & ") AND " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Model0 <> 'Virtual Machine' " & _
	"ORDER BY " & SortBy
	
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>