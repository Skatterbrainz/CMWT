<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adgroup.asp
' lastupdate..... 12/06/2016
' description.... active directory security group information
'-----------------------------------------------------------------------------
time1 = Timer

GroupName = CMWT_GET("gn", "")
QueryOn   = CMWT_GET("qq", "")
PSet      = CMWT_GET("set", "1")
CMWT_VALIDATE GroupName, "No group name specified"

AdsPath = Get_ADsPath (Replace(GroupName, NetSuffix & "\", ""), "group")

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
menulist = "1=General,2=Members,4=Collections"

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
		Response.Write "<table class=""tfx"">"
	
		On Error Resume Next
		Set objGroup = GetObject("LDAP://" & AdsPath)

		Response.Write "" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">SAM Account Name</td>" & _
				"<td class=""td6 v10"">" & objGroup.SAMAccountName & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Canonical Name</td>" & _
				"<td class=""td6 v10"">" & objGroup.Name & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Distinguished Name</td>" & _
				"<td class=""td6 v10"">" & AdsPath & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">E-Mail</td>" & _
				"<td class=""td6 v10"">" & objGroup.Mail & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Information</td>" & _
				"<td class=""td6 v10"">" & objGroup.Info & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Scope</td>" & _
				"<td class=""td6 v10"">"

		If intGroupType AND ADS_GROUP_TYPE_LOCAL_GROUP Then
			Response.Write "Group scope: Domain Local</td></tr>"
		ElseIf intGroupType AND ADS_GROUP_TYPE_GLOBAL_GROUP Then
			Response.Write "Group scope: Global</td></tr>"
		ElseIf intGroupType AND ADS_GROUP_TYPE_UNIVERSAL_GROUP Then
			Response.Write "Group scope: Universal</td></tr>"
		Else
			Response.Write "Group scope: Unknown</td></tr>"
		End If

		If intGroupType AND ADS_GROUP_TYPE_SECURITY_ENABLED Then
			Response.Write "<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Type</td>" & _
				"<td class=""td6 v10"">Security group</td></tr>"
		Else
			Response.Write "<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Type</td>" & _
				"<td class=""td6 v10"">Distribution group</td></tr>"
		End If

		Response.Write "<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Description</td>" & _
				"<td class=""td6 v10"">"
		Select Case VarType(objGroup.Description)
			Case 8:
				Response.write objGroup.Description
			Case Else:
				For Each strValue in objGroup.Description
					Response.Write "<br/>" & Trim(strValue)
				Next
		End Select
		Response.Write "</td></tr>"

		strWhenCreated = objGroup.Get("whenCreated")
		strWhenChanged = objGroup.Get("whenChanged")

		Set objUSNChanged = objGroup.Get("uSNChanged")
		dblUSNChanged = Abs(objUSNChanged.HighPart * 2^32 + objUSNChanged.LowPart)

		Set objUSNCreated = objGroup.Get("uSNCreated")
		dblUSNCreated = Abs(objUSNCreated.HighPart * 2^32 + objUSNCreated.LowPart)

		objGroup.GetInfoEx Array("canonicalName"), 0
		arrCanonicalName = objGroup.GetEx("canonicalName")
		Response.Write "<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Canonical Path</td>" & _
			"<td class=""td6 v10"">"
		For Each strValue in arrCanonicalName
			Response.Write Trim(strValue) & "<br/>"
		Next
		Response.Write "</td></tr>"

		Response.Write "<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">Object class</td>" & _
				"<td class=""td6 v10"">" & objGroup.Class & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">When Created</td>" & _
				"<td class=""td6 v10"">" & strWhenCreated & " (Created - GMT)" & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">When Changed</td>" & _
				"<td class=""td6 v10"">" & strWhenChanged & " (Modified - GMT)" & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">USN Changed</td>" & _
				"<td class=""td6 v10"">" & dblUSNChanged & " (USN Current)" & "</td></tr>" & _
			"<tr class=""tr1""><td class=""td6 v10 w200 bgGray"">USN Created</td>" & _
				"<td class=""td6 v10"">" & dblUSNCreated & " (USN Original)" & "</td></tr>"

		Response.Write "</table>"
	Case "2":
		If CMWT_NotNullString(adspath) Then

			Response.Write "<table class=""tfx"">"
			On Error Resume Next
			Set objGroup = GetObject("LDAP://" & AdsPath)
			objGroup.GetInfo

			arrMemberOf = objGroup.GetEx("member")

			If VarType(arrMemberOf) > 0 Then
				Response.Write "<tr><td class=""td6 v10 bgGray w200"">Name</td>" & _
					"<td class=""td6 v10 bgGray"">Path</td></tr>"
				mcount = 0
				For Each strMember in arrMemberOf
					uid = Get_LogonName(strMember)

					If CMWT_NotNullString(uid) Then
						uid = "<a href=""aduser.asp?uid=" & uid & """ title=""Account Details for: " & uid & """>" & uid & "</a>"
					End If
					Response.Write "<tr class=""tr1""><td class=""td6 v10"">" & uid & "</td>" & _
						"<td class=""td6 v10"">" & strMember & "</td></tr>"
					mcount = mcount + 1
				Next
				Response.Write "<tr class=""tr1""><td class=""td6 v10 bgGray"" colspan=""2"">" & mcount & " members found</td></tr>"
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
		

End Select

CMWT_Footer()
%>

</body>
</html>