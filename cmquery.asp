<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: cmquery.asp
' Author....: David M. Stein
' Date......: 11/27/2016
' Purpose...: configmgr query details
'****************************************************************
time1 = Timer

KeyValue = CMWT_GET("id", "")
SortBy   = CMWT_GET("s", "Name")
QueryON  = CMWT_GET("qq", "")

PageTitle = "Query: " & KeyValue

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT " & _
	"Name,QueryKey,Comments,Architecture," & _
	"Lifetime,QryFmtKey,QueryType, " & _
	"CollectionID,WQL,SQL " & _
	"FROM dbo.Queries " & _
	"WHERE QueryKey = '" & KeyValue & "'"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEROWGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>