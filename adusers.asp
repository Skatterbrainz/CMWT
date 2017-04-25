<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adusers.asp
' lastupdate..... 04/24/2017
' description.... active directory users report
'-----------------------------------------------------------------------------
time1  = Timer
objPfx = CMWT_GET("ch", "A")
SelOpt = CMWT_GET("sel", "")
SortBy = CMWT_GET("s", "sAMAccountName")

PageTitle    = "AD Users"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

blist = "<input type=""button"" id=""bb"" name=""bb"" class=""btx h32 w150"" " & _
	"value=""Disabled"" onClick=""document.location.href='adDisabledUsers.asp'"" />"
CMWT_PageHeading PageTitle, blist

CMWT_CLICKBAR objPfx, "adusers.asp?ch="
	
If objPFX <> "ALL" Then
	query = "<LDAP://" & Application("CMWT_DomainPath") & ">;(&(objectCategory=User)" & _
		"(sAMAccountName=" & objPFX & "*));displayName,userAccountControl,sAMAccountName,whenCreated,pwdLastSet,ADsPath;Subtree"  
Else
	query = "<LDAP://" & Application("CMWT_DomainPath") & ">;(objectCategory=User)" & _
		";displayName,userAccountControl,sAMAccountName,whenCreated,pwdLastSet,ADsPath;Subtree"  
End If

Response.Write "<form name=""form1"" id=""form1"" method=""post"" action=""adUserMod.asp"">" & _
	"<table class=""tfx"">"

On Error Resume Next
Set objConnection = Server.CreateObject("ADODB.Connection")
Set objCommand    = Server.CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Properties("User ID")  = Application("CM_AD_TOOLUSER")
objConnection.Properties("Password") = Application("CM_AD_TOOLPASS")
objConnection.Properties("Encrypt Password") = False
objConnection.Properties("ADSI Flag") = 1
objConnection.Open "Active Directory Provider"
If err.Number <> 0 Then
	Response.Write "[adusers] error-1: " & err.Number & " / " & err.Description
	Response.End
End If

Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 3000
objCommand.CommandText = query

Set objRecordSet = objCommand.Execute
If err.Number <> 0 Then
	Response.Write "[adusers] Error-2: " & err.Number & " / " & err.Description
	Response.WRite "<p>" & Server.URLEncode(query) & "</p>"
	Response.End
End If

If NOT (objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.Sort = SortBy
	objRecordSet.MoveFirst
	xrows = objRecordSet.RecordCount
	
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""td6 v10 w30 bgGray""> </td>" & _
		"<td class=""td6 v10 bgGray"">SAM Account Name</td>" & _
		"<td class=""td6 v10 bgGray"">Display Name</td>" & _
		"<td class=""td6 v10 bgGray"">Status</td>" & _
		"<td class=""td6 v10 bgGray"">ADS Path</td>" & _
		"</tr>"
	
	Do Until objRecordSet.EOF
		udn = objRecordSet.Fields("displayName").value
		sam = objRecordSet.Fields("sAMAccountName").value
		uac = objRecordSet.Fields("userAccountControl").value
		ads = objRecordSet.Fields("ADsPath").value
		adx = Replace(ads, ",", "^")
		samlink = "<a href=""aduser.asp?uid=" & sam & """>" & sam & "</a>"
		uacx = CMWT_UAC(uac)
		
		If Application("CM_AD_TOOLS") = "TRUE" Then
			If SelOpt = "1" Then
				chk = "<input type=""checkbox"" name=""chk"" id=""chk"" value=""" & adx & """ class=""cb1"" checked />"
			Else
				chk = "<input type=""checkbox"" name=""chk"" id=""chk"" value=""" & adx & """ class=""cb1"" />"
			End If
		Else
			chk = ""
		End If
		
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10 w30 ctr"">" & chk & "</td>" & _
			"<td class=""td6 v10"">" & samlink & "</td>" & _
			"<td class=""td6 v10"">" & udn & "</td>" & _
			"<td class=""td6 v10"">" & uacx & "</td>" & _
			"<td class=""td6 v8"">" & ads & "</td>" & _
			"</tr>"
		objRecordSet.MoveNext
	Loop
	
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""5"">" & _
		xrows & " accounts were found</td></tr>"
Else
	Response.Write "<tr class=""tr1 h100""><td class=""td6 v10 ctr"">No matching user accounts were found</td></tr>"
End If

Response.Write "</table>"

If Application("CM_AD_TOOLS") = "TRUE" Then
	Response.Write "<br/>" & _
		"<div class=""t1000x"">" & _
		"<input type=""button"" name=""b0"" id=""b0"" class=""btx w140 h32"" value=""Select All"" onClick=""document.location.href='adusers.asp?ch=" & objPFX & "&sel=1'"" />" & _
		"<input type=""button"" name=""b2"" id=""b2"" class=""btx w140 h32"" value=""Clear All"" onClick=""document.location.href='adusers.asp?ch=" & objPFX & "'"" />" & _
		"<select name=""actiontype"" id=""actiontype"" size=""1"" class=""w200 pad6"">" & _
			"<option value=""""></option>" & _
			"<option value=""DISABLE"">Disable Account</option>" & _
			"<option value=""ENABLE"">Enable Account</option>" & _
		"</select>" & _
		"<input type=""submit"" name=""b1"" id=""b1"" class=""btx w140 h32"" value=""Process"" />"
End If
Response.Write "</form></div>"

CMWT_Footer()
%>

</body>
</html>