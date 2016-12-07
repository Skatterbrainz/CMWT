<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... activeusers.asp
' lastupdate..... 12/03/2016
' description.... current user sessions
'-----------------------------------------------------------------------------
time1 = Timer
IF Not CMWT_ADMIN() Then
	CMWT_STOP "Access Denied!"
End If

PageTitle    = "Active Site Users"
PageBackLink = "admin.asp"
PageBackName = "Administration"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx""><tr><td class=""td6 v10 bgGray"">UserID</td></tr>"

ucount = 0
For each sv in Split(Application("CMWT_USERLIST"),",")
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""td6 v10"">" & sv & "</td>" & _
		"</tr>"
	ucount = ucount + 1
Next

Response.Write "<tr><td class=""td6 v10 bgGray"">" & ucount & " active users found</td></tr></table>"

CMWT_Footer() 
%>
	
</body>
</html>