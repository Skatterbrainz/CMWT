<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... ie.asp
' lastupdate..... 11/27/2016
' description.... microsoft Internet Explorer versions and install counts for each
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer

PageTitle = "IE Version Installs"
SortBy  = CMWT_GET("s", "ProductVersion")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT " & _
	"ProductName, ProductVersion, COUNT(*) AS Installs " & _
	"FROM dbo.v_GS_SoftwareProduct " & _
	"WHERE (ProductName LIKE '%Internet Explorer%') " & _
	"GROUP BY ProductName, ProductVersion " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "ie.asp", "PRODUCTVERSION=ie2.asp?v"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()
%>

</body>
</html>