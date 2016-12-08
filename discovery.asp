<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: discovery.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: site discovery settings summary
'****************************************************************
time1 = Timer

DMName = CMWT_GET("dm","")
CMWT_VALIDATE DMName, "Method Name not provided"

PageTitle = DMName
PageBackLink = "discoveries.asp"
PageBackName = "Discovery Methods"
SortBy  = CMWT_GET("s", "PropertyName")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT " & _
	"dbo.vSMS_SC_Component_Properties.Name AS PropertyName, " & _
	"dbo.vSMS_SC_Component_Properties.Value1, " & _
	"dbo.vSMS_SC_Component_Properties.Value2, " & _
	"dbo.vSMS_SC_Component_Properties.Value3, " & _
	"dbo.vSMS_SC_Component.Flags " & _
"FROM " & _
	"dbo.vSMS_SC_Component INNER JOIN " & _
	"dbo.vSMS_SC_Component_Properties " & _
		"ON dbo.vSMS_SC_Component.ID = dbo.vSMS_SC_Component_Properties.ID " & _
"WHERE " & _
	"(dbo.vSMS_SC_Component.ComponentName = '" & DMName & "') " & _
"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>