<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... ie2.asp
' lastupdate..... 11/27/2016
' description.... computers with specific Internet Explorer version installed
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer

KeyVal = CMWT_GET("v", "")
CMWT_Validate KeyVal, "Version was not specified"
PageTitle = "Computers with IE " & KeyVal
SortBy  = CMWT_GET("s", "DeviceName")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT dbo.v_R_System.Name0 AS DeviceName, " & _
	"dbo.v_GS_SoftwareProduct.ProductName, " & _
	"dbo.v_GS_SoftwareProduct.ProductVersion " & _
	"FROM dbo.v_GS_SoftwareProduct INNER JOIN " & _
    "dbo.v_R_System ON " & _
	"dbo.v_GS_SoftwareProduct.ResourceID = dbo.v_R_System.ResourceID " & _
	"WHERE (dbo.v_GS_SoftwareProduct.ProductName LIKE '%INTERNET EXPLORER%') " & _
	"AND (dbo.v_GS_SoftwareProduct.ProductVersion = '" & KeyVal & "') " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "ie2.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()
%>

</body>
</html>