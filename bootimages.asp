<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... bootimages.asp
' lastupdate..... 12/12/2016
' description.... boot images report
'-----------------------------------------------------------------------------
time1 = Timer
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","Name")

PageTitle    = "Boot Images"
PageBackLink = "software.asp"
PageBackName = "Software"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT DISTINCT " & _
	"PackageID AS PkgID, " & _
	"Name," & _
	"Version," & _
	"Language," & _
	"PkgSourcePath AS SourcePath," & _
	"SourceVersion," & _
	"CONVERT(varchar(10),SourceDate,101)," & _
	"SourceSite " & _
	"FROM dbo.v_BootImagePackage " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>