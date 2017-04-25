<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adgroup.asp
' lastupdate..... 04/24/2017
' description.... active directory security group information
'-----------------------------------------------------------------------------
time1 = Timer

GroupName = CMWT_GET("gn", "")
QueryOn   = CMWT_GET("qq", "")
PSet      = CMWT_GET("set", "1")
CMWT_VALIDATE GroupName, "No group name specified"

AdsPath = CMWT_AD_GetADsPath (Replace(GroupName, NetSuffix & "\", ""), "group")

If Ucase(Application("CM_AD_TOOLS")) = "TRUE" Then
	ad_modify = True
End If
If Ucase(Application("CM_AD_TOOLS_SAFETY")) = "TRUE" Then
	ad_safemode = True
	ad_safelist = Ucase(Application("CM_AD_TOOLS_ADMINGROUPS"))
	If InStr(ad_safelist, Ucase(GroupName)) <> 0 Then
		ad_safelock = True
	End If
End If

PageTitle    = "Group: " & GroupName
PageBackLink = "adgroups.asp"
PageBackName = "Security Groups"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "1=General,2=Members,3=Collections"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	mlink = "adgroup.asp?gn=" & GroupName & "&set=" & mset(0)
	If KeySet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

Select Case PSet
	Case "1":
	
		On Error Resume Next

		fields = "isCriticalSystemObject,whenCreated,mail,groupType,description,distinguishedname,samaccountname"

		query = "SELECT " & fields & " FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
			"WHERE objectClass='group' AND sAMAccountName='" & GroupName & "'"

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

		Response.Write "<table class=""tfx"">"

		If xrows > 0 Then
			Do Until objRecordSet.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to objRecordSet.Fields.Count -1
					fieldname = objRecordSet.Fields(i).Name
					strvalue  = objRecordSet.Fields(i).Value
					Select Case Lcase(fieldname)
						Case "grouptype"
							if strvalue AND ADS_GROUP_TYPE_LOCAL_GROUP then
								strvalue = "Domain Local"
							elseif strvalue AND ADS_GROUP_TYPE_GLOBAL_GROUP then
								strvalue = "Global"
							elseif strvalue AND ADS_GROUP_TYPE_UNIVERSAL_GROUP then
								strvalue = "Universal"
							else
								strvalue = "Unknown"
							end if
						Case "description"
							fv = ""
							for each dv in strValue
								fv = fv & dv
							next
							strvalue = fv
					End Select
					Response.Write "<tr class=""tr1"">" & _
						"<td class=""td6 bgGray v10 w200"">" & CMWT_WordCase(fieldname) & "</td>" & _
						"<td class=""td6 v10"">" & strvalue & "</td></tr>"
				Next
				objRecordSet.MoveNext
			Loop
		else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No matching group found</td></tr>"
		end if

		Response.Write "</table>"

	Case "2":

		' BUG IN THIS SECTION STILL NEEDS TO BE RESOLVED...
		
		If CMWT_NotNullString(adspath) Then
			Response.Write "<table class=""tfx"">"
			On Error Resume Next
			Set objGroup = GetObject("LDAP://" & AdsPath)
			objGroup.GetInfo
			arrMemberOf = objGroup.GetEx("member")
			If VarType(arrMemberOf) > 0 Then
				Response.Write "<tr><td class=""td6 v10 bgGray"">User ID</td>" & _
					"<td class=""td6 v10 bgGray w200"">Name</td>" & _
					"<td class=""td6 v10 bgGray"">Path</td></tr>"
				mcount = 0
				For Each strMember in arrMemberOf
					'uid = Get_LogonName(strMember)
					'uid = CMWT_AD_SamAccountName("LDAP://" & strMember, "user")
					uid = Replace(Split(strMember,",")(0), "CN=","")
					set openDS  = GetObject("LDAP:")
					if err.number <> 0 Then response.write "exception1: " & err.Description : response.end

					set objUser  = openDS.OpenDSObject("LDAP://" & strMember, Application("CM_AD_TOOLUSER"), Application("CM_AD_TOOLPASS"), ADS_SECURE_AUTHENTICATION)
					if err.number <> 0 Then response.write "exception2: " & err.Description : response.end
					uid = objUser.samaccountname
					ucn = objUser.displayName
					
					If CMWT_NotNullString(uid) Then
						uid = "<a href=""aduser.asp?uid=" & uid & """ title=""Account Details for: " & uid & """>" & uid & "</a>"
					End If
					Response.Write "<tr class=""tr1""><td class=""td6 v10"">" & uid & "</td>" & _
						"<td class=""td6 v10"">" & ucn & "</td>" & _
						"<td class=""td6 v10"">" & strMember & "</td></tr>"
					mcount = mcount + 1
				Next
				Response.Write "<tr class=""tr1""><td class=""td6 v10 bgGray"" colspan=""3"">" & mcount & " members found</td></tr>"
			Else
				Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No members found</td></tr>"
			End If
			Response.Write "</table>"
			If ad_modify = True Then
				If ad_safelock = True Then 
					Response.Write "<div class=""tfx v10""><p>" & _
						"This group is protected from CMWT change requests.</p></div>"
				Else
					Response.Write "<div class=""tfx""><br/>" & _
						"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" value=""Add Member"" />" & _
						"</div>"
				End If
			End If
		End If
		
	Case "3":
		
		query = "SELECT DISTINCT dbo.v_Collection.CollectionID, dbo.v_Collection.Name, " & _
			"CASE WHEN dbo.v_Collection.CollectionType=1 THEN 'USER' ELSE 'DEVICE' END AS [CollectionType], dbo.v_Collection.Comment " & _
			"FROM dbo.v_Collection INNER JOIN " & _
			"dbo.v_CollectionRuleQuery ON dbo.v_Collection.CollectionID = dbo.v_CollectionRuleQuery.CollectionID " & _
			"WHERE (dbo.v_CollectionRuleQuery.QueryExpression LIKE '%" & GroupName & "%') " & _
			"ORDER BY v_Collection.Name"

		Dim conn, cmd, rs
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEGRID rs, "", "", ""
		CMWT_DB_CLOSE()
		CMWT_SHOW_QUERY()

End Select

CMWT_Footer()
%>

</body>
</html>