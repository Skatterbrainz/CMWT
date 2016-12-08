<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmusers.asp
' lastupdate..... 12/04/2016
' description.... sccm user accounts
'-----------------------------------------------------------------------------
time1 = Timer

UserID = CMWT_GET("uid", "")
objPfx = CMWT_GET("ch", "A")
SortBy = CMWT_GET("s", "User_Name0")
ADDom  = CMWT_GET("d", "")
RMode  = CMWT_GET("x", "")

If CMWT_NotNullString(UserID) Then
	PageTitle = "User Accounts: " & UserID
ElseIf RMode = "1" Then
	PageTitle = "User Accounts by Domain"
	SortBy = "DOMAIN"
Else
	PageTitle = "User Accounts"
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "cmusers.asp?ch="

Response.Write "<table class=""tfx"">"
'----------------------------------------------------------------
fields = "Full_User_Name0 AS DisplayName, User_Name0 AS UserID, Full_Domain_Name0 AS DomainName, Mail0 AS Email"

If objPFX <> "ALL" Then
	query = "SELECT DISTINCT " & fields & " " & _
		"FROM dbo.v_R_User " & _
		"WHERE User_Name0 LIKE '" & objPFX & "%' " & _
		"AND Full_Domain_Name0='" & Application("CMWT_DOMAINSUFFIX") & "' " & _
		"ORDER BY " & SortBy
Else
	query = "SELECT DISTINCT " & fields & " " & _
		"FROM dbo.v_R_User " & _
		"WHERE Full_Domain_Name0='" & Application("CMWT_DOMAINSUFFIX") & "' " & _
		"ORDER BY " & SortBy
End If

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query

If Not(rs.BOF And rs.EOF) Then
	found = True
	xrows = rs.RecordCount
	xcols = rs.Fields.Count
	
	Response.Write "<tr>"
	For i = 0 to xcols-1
		fn = rs.Fields(i).Name
		Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
	Next
	Response.Write "</tr>"
	
	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			Select Case Ucase(fn)
				Case "MAIL","EMAIL","MAIL0":
					fv = "<a href=""mailto:" & fv & """ title=""Send Email to " & fv & """>" & fv & "</a>"
				Case "DOMAIN","DOMAINNAME":
					udom = fv
					fv = "<a href=""cmusers.asp?d=" & fv & """ title=""Filter on domain " & fv & """>" & fv & "</a>"
				Case "USER_NAME0","USERID":
					fv = "<a href=""cmuser.asp?uid=" & fv & """>" & fv & "</a>"
			End Select
			Response.Write "<td class=""td6 v10"">" & fv & "</td>"
		Next
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
		xrows & " rows returned</td></tr>"
Else
	Response.Write "<tr class=""h100 tr1"">" & _
		"<td class=""td6 v10 ctr"">No matching items returned</td></tr>"
End If

CMWT_DB_CLOSE()
Response.Write "</table>"
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>
