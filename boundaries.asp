<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... boundaries.asp
' lastupdate..... 11/30/2016
' description.... site boundaries report
'-----------------------------------------------------------------------------
time1 = Timer

SortBy    = CMWT_GET("s", "DisplayName")
QueryOn   = CMWT_GET("qq", "")
PageTitle = "Site Boundaries"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT DISTINCT BoundaryID, DisplayName, " & _
	"CASE WHEN BoundaryType=1 THEN 'Site' " & _
	"WHEN BoundaryType=3 THEN 'IP Range' " & _
	"ELSE '' END AS BoundaryTypeName, " & _
	"Value, GroupCount " & _
	"FROM dbo.vSMS_Boundary " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "boundaries.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>