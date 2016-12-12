<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dpgroups.asp
' lastupdate..... 12/09/2016
' description.... distribution point groups report
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "Name")
QueryOn = CMWT_GET("qq", "")
PageTitle    = "Distribution Point Groups"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%


query = "SELECT Name AS DPGroup,Description,MembersCount " & _
	"FROM dbo.vSMS_DPGroupInfo " & _
	"ORDER BY " & SortBy
	
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "dpgroups.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>
