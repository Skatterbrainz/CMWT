<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... mfrs.asp
' lastupdate..... 11/30/2016
' description.... device inventory counts by vendor
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle = "Devices by Manufacturer"
PageBackLink = "assets.asp"
PageBackName = "Assets"
SortBy  = CMWT_GET("s", "Manufacturer")
QueryOn = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT Manufacturer, COUNT(*) AS QTY " & _
	"FROM (" & _
	"SELECT DISTINCT " & _
		"dbo.v_R_System.Name0 AS DeviceName, " & _
		"COALESCE (dbo.v_GS_COMPUTER_SYSTEM.Manufacturer0, 'UNKNOWN') AS Manufacturer, " & _
		"dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model, " & _
		"dbo.v_R_System.AD_Site_Name0 AS ADSiteName, " & _
		"dbo.v_R_System.Client_Version0 AS ClientVersion, " & _
		"COALESCE(dbo.v_GS_COMPUTER_SYSTEM.SystemType0, '') AS CpuType " & _
	"FROM dbo.v_R_System LEFT OUTER JOIN dbo.v_GS_COMPUTER_SYSTEM ON " & _
		"dbo.v_R_System.ResourceID = dbo.v_GS_COMPUTER_SYSTEM.ResourceID " & _
	") AS T1 " & _
	"GROUP BY T1.Manufacturer " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "mfrs.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()
%>
	
</body>
</html>