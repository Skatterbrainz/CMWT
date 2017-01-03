<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... osimages.asp
' lastupdate..... 01/02/2017
' description.... operating system images
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

PageTitle    = "OS Images"
PageBackLink = "software.asp"
PageBackName = "Software"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT PackageID, " & _
	"Name, " & _
	"OSVersion, " & _
	"Language, " & _
	"ProductType, " & _
	"ROUND(Size/1024/1024/1024,8,3) AS [SIZE] " & _
	"FROM dbo.vSMS_ImageInformation " & _
	"WHERE Name NOT LIKE '%Windows PE%' " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "osimages.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
