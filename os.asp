<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... os.asp
' lastupdate..... 12/03/2016
' description.... operating systems inventory summary
'-----------------------------------------------------------------------------
time1 = Timer

OSName  = CMWT_GET("on", "")
SortBy  = CMWT_GET("s", "Name0")
QueryON = CMWT_GET("qq", "")

CMWT_VALIDATE OSName, "Operating System Name was not specified"
PageTitle    = OSName
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT Name0 AS ComputerName, " & _
	"Client0 AS Client, " & _
	"AD_Site_Name0 AS ADSiteName, " & _
	"Manufacturer0 AS Manufacturer, " & _
	"Model0 AS ModelName, " & _
	"TotalPhysicalMemory0 AS Memory, " & _
	"SystemType0 AS CPUType " & _
	"FROM (" & q_devices & ") AS T1 " & _
	"WHERE T1.Caption0='" & OSName & "' " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "os.asp?on=" & OSName, ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()
%>

</body>
</html>