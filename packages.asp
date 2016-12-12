<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... packages.asp
' lastupdate..... 11/27/2016
' description.... package library
'-----------------------------------------------------------------------------
time1 = Timer
objPfx  = CMWT_GET("ch", "ALL")
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","Name")

PageTitle = "Packages"
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "packages.asp?ch="

If objPFX <> "ALL" Then
	query = "SELECT Name, PackageID, Version, Manufacturer, SourceVersion " & _
		"FROM dbo.v_Package " & _
		"WHERE (Name LIKE '" & objPfx & "%') " & _
		"AND (PackageType=0) " & _
		"ORDER BY " & SortBy
Else
	query = "SELECT Name, PackageID, Version, Manufacturer, SourceVersion " & _
		"FROM dbo.v_Package " & _
		"WHERE PackageType=0 " & _
		"ORDER BY " & SortBy
End If

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "packages.asp", "PackageID~package.asp?id="
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>