<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... depsummary2.asp
' lastupdate..... 01/02/2017
' description.... app deployment summary: short view
'-----------------------------------------------------------------------------
time1 = Timer

SortBy   = CMWT_GET("s", "DisplayName")
FilterFN = CMWT_GET("fn", "")
FilterFV = CMWT_GET("fv", "")
QueryON  = CMWT_GET("qq", "")

if FilterFV <> "" Then
	PageTitle    = "Deployment Summary: " & FilterFV
else
	PageTitle    = "Deployment Summary"
end if
PageBackLink = "software.asp"
PageBackName = "Software"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT " & _
	"DisplayName," & _
	"AppCI," & _
	"DevicesWithApp," & _
	"UsersTargetedWithApp," & _
	"DevicesWithFailure," & _
	"UsersWithFailure," & _
	"UsersRequested " & _
	"FROM dbo.vAppStatSummary "
If FilterFN <> "" Then
	query = query & _
		" WHERE (" & FilterFN & " = '" & FilterFV & "') " & _
		"ORDER BY " & SortBy
	filtered = True
Else
	query = query & " ORDER BY " & SortBy	
End If

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TableGridFilter rs, "", "depsummary2.asp", "", "", "depsummary2.asp"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
