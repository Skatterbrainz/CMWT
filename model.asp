<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... model.asp
' lastupdate..... 11/30/2016
' description.... computers by specified model name
'-----------------------------------------------------------------------------
time1 = Timer

mn = CMWT_GET("m", "")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE mn, "No model name was provided"

PageTitle = mn
PageBackLink = "models.asp"
PageBackName = "Computer Models"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
		"Name0 AS ComputerName, " & _
		"ResourceID, " & _
		"AD_Site_Name0 AS ADSiteName, " & _
		"SystemType0 AS SystemType, " & _
		"Client0 AS Client, " & _
		"Caption0 AS WindowsVersion, " & _
		"Full_Domain_Name0 AS DomainName, " & _
		"UserName0 AS UserName " & _
	"FROM (" & q_devices & ") AS T1 " & _
	"WHERE T1.Model0='" & mn & "'"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "model.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>
	
</body>
</html>