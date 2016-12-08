<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... products.asp
' lastupdate..... 12/04/2016
' description.... installed software applications inventory
'-----------------------------------------------------------------------------
time1 = Timer
objPfx  = CMWT_GET("ch", "A")
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","ProductName")

PageTitle = "Installed Software"
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "products.asp?ch="

If objPFX <> "ALL" Then
	query = "SELECT DISTINCT " & _
		"ARPDisplayName0 AS ProductName, " & _
		"NormalizedPublisher AS Publisher, " & _
		"ProductCode0 AS ProductCode, COUNT(ResourceID) AS Installs " & _
		"FROM dbo.v_GS_INSTALLED_SOFTWARE_CATEGORIZED " & _
		"WHERE (ARPDisplayName0 LIKE '" & objPfx & "%') " & _
		"GROUP BY ARPDisplayName0, NormalizedPublisher, ProductCode0 " & _
		"ORDER BY " & SortBy
Else
	query = "SELECT DISTINCT " & _
		"ARPDisplayName0 AS ProductName, " & _
		"NormalizedPublisher AS Publisher, " & _
		"ProductCode0 AS ProductCode, COUNT(ResourceID) AS Installs " & _
		"FROM dbo.v_GS_INSTALLED_SOFTWARE_CATEGORIZED " & _
		"WHERE (ARPDisplayName0 IS NOT NULL) AND (ARPDisplayName0 <> '') " & _
		"GROUP BY ARPDisplayName0, NormalizedPublisher, ProductCode0 " & _
		"ORDER BY " & SortBy
End If

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "products.asp?ch=" & objPfx, ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>