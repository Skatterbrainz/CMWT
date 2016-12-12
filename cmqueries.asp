<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: cmqueries.asp
' Author....: David M. Stein
' Date......: 12/10/2016
' Purpose...: configmgr queries
'****************************************************************
time1 = Timer
SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

PageTitle    = "Queries"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT " & _
	"QueryID, Name, Comments, TargetClassName " & _
	"FROM dbo.v_Query " & _
	"ORDER BY " & SortBY
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "cmqueries.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
Response.Write "</body></html>"
%>
