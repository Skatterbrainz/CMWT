<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: discoveries.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: site discovery settings summary
'****************************************************************
time1 = Timer

PageTitle = "Discovery Methods"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "Discovery")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT ComponentName AS Discovery " & _
	"FROM dbo.vSMS_SC_Component " & _
	"WHERE ComponentName LIKE '%DISC%' " & _
	"ORDER BY " & SortBY

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", "DISCOVERY=discovery.asp?dm="
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>