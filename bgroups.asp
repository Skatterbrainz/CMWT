<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... bgroups.asp
' lastupdate..... 11/30/2016
' description.... site boundary groups
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "Boundary Groups"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "BoundaryGroup")
QueryOn = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT Name AS BoundaryGroup,GroupID,Description,DefaultSiteCode," & _ 
"MemberCount,SiteSystemCount,Shared " & _
"FROM dbo.vSMS_BoundaryGroup " & _
"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "bgroups.asp", "NAME^bdp.asp?gn="
CMWT_DB_CLOSE()
CMWT_SHOW_Query()
CMWT_Footer()
%>

</body>
</html>