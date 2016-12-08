<!-- #include file="_core.asp" -->
<%
'****************************************************************
' Filename..: diag.asp
' Author....: David M. Stein
' Date......: 12/03/2016
' Purpose...: application diagnostics information
'****************************************************************
time1 = Timer
Response.Expires = -1

If Not CMWT_ADMIN() Then
	Response.Redirect "error.asp?m=Access Denied / Unauthorized User"
End If

PageTitle    = "Diagnostics"
PageBackLink = "admin.asp"
PageBackName = "Administration"

SortBy  = CMWT_GET("s", "ServerName")
QueryON = CMWT_GET("qq", "")
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<tr>
			<td class="td6 v10 bgGray">Variable</td>
			<td class="td6 v10 bgGray">Assigned Value</td>
		</tr>
		<%
		For each sv in Session.Contents
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10"">" & sv & "</td>" & _
				"<td class=""td6 v10"">" & Session(sv) & "</td>" & _
				"</tr>"
		Next
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10"">BROWSER TYPE</td>" & _
			"<td class=""td6 v10"">" & CMWT_BROWSER_TYPE() & "</td>" & _
			"</tr>"
		%>
	</table>
	
	<% 
	'Response.End
	IF CMWT_ADMIN() Then 
	%>

	<br/>
	<div class="t1000x"><h3>Application Data</h3></div>
	
	<table class="t1000x">
		<tr>
			<td class="td6 v10 bgGray">Variable</td>
			<td class="td6 v10 bgGray">Assigned Value</td>
		</tr>
		<%
		For each sv in Application.Contents
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10"">" & sv & "</td>" & _
				"<td class=""td6 v10"">" & Application(sv) & "</td>" & _
				"</tr>"
		Next
		%>
	</table>
	
	<div class="t1000x"><h3>Server Configuration</h3></div>
	
	<table class="t1000x">
		<tr>
			<td class="td6 v10 bgGray">Variable</td>
			<td class="td6 v10 bgGray">Assigned Value</td>
		</tr>
		<%
		For each sv in Request.ServerVariables()
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10"">" & sv & "</td>" & _
				"<td class=""td6 v10"">" & Request.ServerVariables(sv) & "</td>" & _
				"</tr>"
		Next
		%>
	</table>
	<% 
	End If
	CMWT_Footer() 
	%>
	
</body>
</html>