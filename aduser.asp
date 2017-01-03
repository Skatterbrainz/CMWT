<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... aduser.asp
' lastupdate..... 01/02/2017
' description.... active directory user account information
'-----------------------------------------------------------------------------
time1 = Timer

UserID  = CMWT_GET("uid", "")
Pset    = CMWT_GET("set", "1")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE UserID, "User Account ID was not specified"

If InStr(UserID, "\") > 0 Then
	uu = Split(UserID,"\")
	UserID = uu(1)
End If

If PSet = "4" Then
	Response.Redirect "cmuser.asp?uid=" & UserID
End If

PageTitle    = UserID 
PageBackLink = "adusers.asp"
PageBackName = "AD Users"

On Error Resume Next

fields = "logonCount,userAccountControl,pwdLastSet," & _
	"whenCreated,userWorkstations,employeeNumber,employeeID,manager,st,l," & _
	"streetAddress,physicalDeliveryOfficeName,telephoneNumber," & _
	"ipPhone,mobile,facsimileTelephoneNumber,department,company,mail," & _
	"description,title,samaccountname,displayName"

query = "SELECT " & fields & " FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
	"WHERE objectClass='user' AND sAMAccountName='" & UserID & "'"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "1=Account,2=Groups,3=Computers"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	'aduser.asp?uid=" & UserID & "&set=2
	mlink = "aduser.asp?uid=" & UserID & "&set=" & mset(0)
	If KeySet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

Select Case Pset

	Case "1":
	
		Response.Write "<table class=""tfx"">"

		arrFN = Split(fields,",")
		xcols = Ubound(arrFN)
		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand    = CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID")  = Application("CM_AD_TOOLUSER")
		objConnection.Properties("Password") = Application("CM_AD_TOOLPASS")
		objConnection.Properties("ADSI Flag") = 1
		objConnection.Open "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection
		objCommand.Properties("Page Size") = 1000
		objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
		objCommand.CommandText = query
		Set objRecordSet = objCommand.Execute
		objRecordSet.MoveFirst
		xrows = objRecordSet.RecordCount

		If xrows > 0 Then
			Do Until objRecordSet.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to objRecordSet.Fields.Count -1
					fieldname = objRecordSet.Fields(i).Name
					strvalue  = objRecordSet.Fields(i).Value
					fn = fieldname
					Select Case Ucase(fieldname)
						Case "NAME":
							'fn = "Name"
							fv = strValue
							funx = CMWT_NameParse (cn)
						Case "DESCRIPTION":
							'fn = "Description"
							If CMWT_NotNullString(strValue) Then
								d = ""
								For each x in strValue
									d = d & x
								Next
								fv = d
							Else
								fv = ""
							End If
						Case "USERACCOUNTCONTROL":
							fv = CMWT_UAC (fv)
							fn = "Status"
						Case "L","CITY":
							fn = "City"
							fv = strValue
						Case "ST":
							fn = "State"
							fv = strValue
						Case "FACSIMILETELEPHONENUMBER":
							fn = "FAX"
							fv = strValue
						Case Else:
							fv = strValue
							fn = fieldname
					End Select

					Response.Write "<tr class=""tr1"">" & _
						"<td class=""td6 v10 bgGray w200"">" & CMWT_WordCase(fn) & "</td>" & _
						"<td class=""td6 v10"">" & fv & "</td></tr>"
				Next
				Response.Write "</tr>"
				objRecordSet.MoveNext
			Loop

		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No matching account found</td></tr>"
		End If
		
		Response.Write "</table>"
			
	Case "2":
	
		Response.Write "<table class=""tfx""><tr><td class=""td6 v10 bgGray"">Group Name</td></tr>"
		groups = CMWT_AD_EnumGroups (UserID)
		If CMWT_NotNullString(groups) Then
			gcount = 0
			For each group in Split(groups, ",")
				glink = "adgroup.asp?gn=" & group
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10 ptr"" onClick=""document.location.href='" & glink & _
					"'"" title=""View Group Details..."">" & group & "</td></tr>"
				gcount = gcount + 1
			Next
		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No group memberships found</td></tr>"
		End If
		Response.Write "</table>"

		If ADModify = True Then
			Response.Write "<div class=""tfx""><br/>" & _
				"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" value=""Add to Group"" />" & _
				"</div>"
		Else
			Response.Write "<div class=""tfx""><br/>" & _
				"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" value=""Add to Group"" disabled=""true"" />" & _
				"</div>"
		End If
	Case "3":
		Dim conn, cmd, rs
		query = "SELECT DISTINCT " & _
			"dbo.v_R_System.Name0 AS ComputerName, dbo.v_R_System.AD_Site_Name0 AS ADSiteName, " & _
			"dbo.v_GS_USER_PROFILE.LocalPath0 AS ProfilePath, dbo.v_GS_USER_PROFILE.TimeStamp " & _
			"FROM dbo.v_GS_USER_PROFILE INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_USER_PROFILE.ResourceID = dbo.v_R_System.ResourceID INNER JOIN " & _
			"dbo.v_R_User ON dbo.v_GS_USER_PROFILE.SID0 = dbo.v_R_User.SID0 " & _
			"WHERE (dbo.v_R_User.User_Name0 = '" & UserID & "')"

		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEGRID rs, "", "", ""
		CMWT_DB_CLOSE()
		
End Select

CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>