<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... vmhosts.asp
' lastupdate..... 12/04/2016
' description.... virtual machine hosts
'-----------------------------------------------------------------------------
time1 = Timer
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","VMHost")

PageTitle = "Virtual Hosts"
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT Virtual_Machine_Host_Name0 AS VMHOST, COUNT(ResourceID) AS Guests " & _
	"FROM dbo.v_R_System " & _
	"WHERE (Virtual_Machine_Host_Name0 IS NOT NULL) AND (LTRIM(Virtual_Machine_Host_Name0) <> '') " & _
	"GROUP BY Virtual_Machine_Host_Name0 " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "vmhosts.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>