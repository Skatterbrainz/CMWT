<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... oulist.asp
' lastupdate..... 11/30/2016
' description.... interactive AD OU browser
'-----------------------------------------------------------------------------
time1 = Timer

oupath  = CMWT_GET("ou", "")
objType = CMWT_GET("type", "user")
findFN  = CMWT_GET("f", "")
findFV  = CMWT_GET("v", "")

PageTitle = "AD Organizational Units"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

If ouPath <> "" Then
	LINKPATH = "accounts.asp?ou=" & oupath
Else
	LINKPATH = ""
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
	
<form name="form2" id="form2" method="post" action="">
	<table class="tfx">
		<tr>
			<td class="v10 vtop w300">
				<select name="ou" id="ou" size="28" class="w300 pad6 bgGray" onChange="if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'frame1') }">
					<% xrows = EnumerateOUs() %>
				</select><br/>
				<%=xrows%> OUs were returned
			</td>
			<td class="v10 vtop">
				<iframe name="frame1" id="frame1" style="width:880px;height:500px;" frameborder="0" src="<%=LINKPATH%>">Unsupported Web Browser</iframe>
			</td>
		</tr>
	</table>
</form>

<% CMWT_Footer() %>
	
</body>
</html>