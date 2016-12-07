<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adusermod.asp
' lastupdate..... 12/04/2016
' description.... modify AD user account properties
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer

PageTitle = "Manage AD User Accounts"
ActType   = CMWT_GET("actiontype", "")
Accounts  = CMWT_GET("chk", "")
GroupName = CMWT_GET("gn", "")

CMWT_VALIDATE Accounts, "No accounts were selected"
CMWT_VALIDATE ActType, "Action Type Code was not provided"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<%

blist = "<input type=""button"" id=""bb"" name=""bb"" class=""btx h32 w150"" " & _
	"value=""Return"" onClick=""document.location.href='adUsers.asp'"" />"
CMWT_PageHeading PageTitle, blist

count_1 = 0
count_2 = 0
count_3 = 0
count_X = 0

Response.Write "<table class=""tfx"">" & _
	"<tr><td class=""td6 v10 bgGray"">Account</td>" & _
	"<td class=""td6 v10 bgGray"">Action</td>" & _
	"<td class=""td6 v10 bgGray"">Result</td></tr>"

If Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
	Dim conn, cmd, rs
	CMWT_DB_OPEN Application("DSN_CMWT")
End If

For each uid in Split(Accounts, ",")
	uid = Replace(Trim(uid), "^", ",")
	on error resume next
	set openDS = GetObject("LDAP:")
	set objUser = openDS.OpenDSObject(uid, Application("CM_AD_TOOLUSER"), Application("CM_AD_TOOLPASS"), ADS_SECURE_AUTHENTICATION)
	if err.number = 0 then
		select case Ucase(ActType)
			case "ENABLE":
				if objUser.AccountDisabled = True then
					objUser.AccountDisabled = FALSE
					objUser.SetInfo
					if err.Number = 0 then
						count_1 = count_1 + 1
						result = "Account has been enabled"
						CMWT_LogEvent conn, "AD", "USER ENABLE", uid & " has been enabled by " & CMWT_USERNAME()
					else
						count_2 = count_2 + 1
						result = "Error: Account was not enabled"
					end if
				else
					count_3 = count_3 + 1
					result = "Account was already enabled"
				end if
			case "DISABLE":
				if objUser.AccountDisabled = False then
					objUser.AccountDisabled = TRUE
					objUser.SetInfo
					if err.Number = 0 then
						count_1 = count_1 + 1
						result = "Account has been disabled"
						CMWT_LogEvent conn, "AD", "USER DISABLE", uid & " has been disabled by " & CMWT_USERNAME()
					else
						count_2 = count_2 + 1
						result = "Error: Account was not disabled"
					end if
				else
					count_3 = count_3 + 1
					result = "Account was already disabled"
				end if
		end select
	else
		result = err.Number
		Response.Write "<br/>your shit blew up. (" & result & ")"
	end if
	set objUser = Nothing
	err.Clear
	count_x = count_x + 1
	Response.Write "<tr class=""tr1""><td class=""td6 v10"">" & uid & "</td>" & _
		"<td class=""td6 v10"">" & ActType & "</td>" & _
		"<td class=""td6 v10"">" & result & "</td></tr>"
Next

If Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
	CMWT_DB_CLOSE()
End If

Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""3"">" & _
	count_x & " accounts processed: Success=" & count_1 & ", Failed: " & count_2 & ", Skipped=" & count_3 & _
	"</td></tr></table>"

CMWT_Footer()
%>

</body>
</html>