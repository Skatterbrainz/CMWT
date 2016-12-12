<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... mydevice.asp
' lastupdate..... 12/10/2016
' description.... my device information
'-----------------------------------------------------------------------------
time1 = Timer

cn = CMWT_GET("cn", "")
If cn = "" Then
	ipa = Trim(Request.ServerVariables("REMOTE_ADDR"))
	If ipa <> "" Then
		Dim conn
		Set conn = Server.CreateObject("ADODB.Connection")
		On Error Resume Next
		conn.ConnectionTimeOut = 5
		conn.Open Application("DSN_CMDB")
		If err.Number <> 0 Then
			CMWT_STOP "error: database connection failure"
		End If
		cn = CMWT_HOST_BY_IP (c, ipa)
		If cn <> "" Then
			cn = cn & "<br/>(based on IP address " & ipa & ")"
		End If
		conn.Close
	Else
		reason = "Unable to capture remote IPv4 address via browser session."
	End If
End If

pageTitle = "My Device"

CMWT_NewPage "", "", ""
PageBackLink = "assets.asp"
PageBackName = "Assets"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx""><tr class=""h300""><td class=""td6 v10 ctr"">"

If cn <> "" Then 
	acn = Split(cn, ".")
	hostname = acn(0)
	golink = "device.asp?cn=" & hostname
	
	Response.Write "<h2>" & cn & "</h2>" & _
		"<p><input type=""button"" name=""Btn1"" id=""Btn1"" class=""w140 h32 btx"" " & _
		"value=""More..."" onClick=""document.location.href='" & golink & "'"" " & _
		"title=""View more device details"" /></p>"
Else
	Response.Write "Sorry. We are unable to identify DNS name for IP address <strong>" & ipa & "</strong>" & _
		"<br/><br/>Tip: Create a Shortcut to ""http://" & SiteServer & "/cmwt/mydevice.asp?cn=%COMPUTERNAME%""" & _
		"<br/>Name it 'My Device' and use that shortcut to launch the web page for a direct look-up."
End If
Response.Write "</td></tr></table>"

CMWT_Footer() 

%>
	
</body>
</html>