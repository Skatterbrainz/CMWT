<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adDisabledUsers.asp
' lastupdate..... 12/04/2016
' description.... disabled active directory user accounts
'-----------------------------------------------------------------------------
time1 = Timer
Response.Expires = -1

SelOpt = CMWT_GET("sel","")

PageTitle = "Disabled AD Users"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

If Ucase(Application("CM_AD_TOOLS")) = "TRUE" Then
	enable_tools = True
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

blist = "<input type=""button"" id=""bb"" name=""bb"" class=""btx h32 w150"" " & _
	"value=""All Users"" onClick=""document.location.href='adUsers.asp'"" />"

CMWT_PageHeading PageTitle, blist

Response.Write "<form name=""form1"" id=""form1"" method=""post"" action=""adUserMod.asp"">" & _
	"<table class=""tfx"">"

On Error Resume Next
Set objConnection = Server.CreateObject("ADODB.Connection")
Set objCommand = Server.CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000

objCommand.CommandText = _
	"<LDAP://" & Application("CMWT_DOMAINPATH") & ">;(&(objectCategory=User)" & _
		"(userAccountControl:1.2.840.113556.1.4.803:=2));Name,sAMAccountName,ADsPath;Subtree"  
Set objRecordSet = objCommand.Execute

If NOT (objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.MoveFirst
	xrows = objRecordSet.RecordCount
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""td6 v10 w30 bgGray"">&nbsp;</td>" & _
		"<td class=""td6 v10 bgGray"">Name</td>" & _
		"<td class=""td6 v10 bgGray"">sAMAccountName</td>" & _
		"<td class=""td6 v10 bgGray"">ADS Path</td>" & _
		"</tr>"
	Do Until objRecordSet.EOF
		udn  = objRecordSet.Fields("Name").Value
		sam  = objRecordSet.Fields("sAMAccountName").value
		ads  = objRecordSet.Fields("ADsPath").value
		adx  = Replace(ads, ",", "^")
		If enable_tools = True Then 
			If SelOpt = "1" Then
				chk = "<input type=""checkbox"" name=""chk"" id=""chk"" value=""" & adx & """ class=""cb1"" checked />"
			Else
				chk = "<input type=""checkbox"" name=""chk"" id=""chk"" value=""" & adx & """ class=""cb1"" />"
			End If
		Else
			chk = "<input type=""checkbox"" name=""chk"" id=""chk"" value="""" class=""cb1"" disabled />"
		End If
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10 w30 ctr"">" & chk & "</td>" & _
			"<td class=""td6 v10"">" & udn & "</td>" & _
			"<td class=""td6 v10"">" & sam & "</td>" & _
			"<td class=""td6 v8"">" & ads & "</td>" & _
			"</tr>"
		objRecordSet.MoveNext
	Loop
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""4"">" & _
		xrows & " accounts were found</td></tr>"
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No disabled user accounts were found</td></tr>"
End If

Response.Write "</table>"

If Ucase(Application("CM_AD_TOOLS")) = "TRUE" Then
	Response.Write "<div class=""tfx"">" & _
	"<input type=""button"" name=""b0"" id=""b0"" class=""btx w140 h32"" value=""Select All"" onClick=""document.location.href='adDisabledUsers.asp?sel=1'"" />" & _
	"<input type=""button"" name=""b2"" id=""b2"" class=""btx w140 h32"" value=""Clear All"" onClick=""document.location.href='adDisabledUsers.asp'"" />" & _
	"<select name=""actiontype"" id=""actiontype"" size=""1"" class=""w200 pad6"">" & _
		"<option value=""""></option>" & _
		"<option value=""ENABLE"">Enable Account</option>" & _
	"</select>" & _
	"<input type=""submit"" name=""b1"" id=""b1"" class=""btx w140 h32"" value=""Process"" />" & _
	"</form></div>"
End If

CMWT_Footer()
Response.Write "</body></html>"
%>
