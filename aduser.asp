<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... aduser.asp
' lastupdate..... 11/30/2016
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

PageTitle = "User Account: " & UserID 
PageBackLink = "adusers.asp"
PageBackName = "AD Users"

On Error Resume Next

fields = "logonCount,userAccountControl,pwdLastSet," & _
	"whenCreated,employeeNumber,employeeID,manager,st,l," & _
	"streetAddress,physicalDeliveryOfficeName,telephoneNumber," & _
	"ipPhone,mobile,facsimileTelephoneNumber,department,company,mail," & _
	"description,samaccountname,displayName"

query = "SELECT " & fields & " FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
	"WHERE objectClass='user' AND sAMAccountName='" & UserID & "'"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tfx">
	<tr>
		<td class="pad6 a14 w300"><strong><%=UserID%></strong></td>
		<td> </td>
		<%
		Select Case Pset
			Case "1":
				Response.Write "<td class=""pad5a v10 w200 bgBlue ctr"">Account</td>" & _
					"<td class=""pad5a v10 w200 ctr ptr"" onMouseOver=""this.className='pad5a v10 w200 ctr ptr bgGray'"" onMouseOut=""this.className='pad5a v10 w200 ctr ptr'"" onClick=""document.location.href='aduser.asp?uid=" & UserID & "&set=2'"">Groups</td>" & _
					"<td class=""pad5a v10 w200 ctr ptr"" onMouseOver=""this.className='pad5a v10 w200 ctr ptr bgGray'"" onMouseOut=""this.className='pad5a v10 w200 ctr ptr'"" onClick=""document.location.href='cmuser.asp?uid=" & UserID & "'"">CM Account</td>"
			Case "2":
				Response.Write "<td class=""pad5a v10 w200 ctr ptr"" onMouseOver=""this.className='pad5a v10 w200 ctr ptr bgGray'"" onMouseOut=""this.className='pad5a v10 w200 ctr ptr'"" onClick=""document.location.href='aduser.asp?uid=" & UserID & "&set=1'"">Account</td>" & _
					"<td class=""pad5a v10 w200 bgBlue ctr"">Groups</td>" & _
					"<td class=""pad5a v10 w200 ctr ptr"" onMouseOver=""this.className='pad5a v10 w200 ctr ptr bgGray'"" onMouseOut=""this.className='pad5a v10 w200 ctr ptr'"" onClick=""document.location.href='cmuser.asp?uid=" & UserID & "'"">CM Account</td>"
		End Select
		%>
	</tr>
</table>
	
<%
Select Case Pset

	Case "1":
	
		Response.Write "<table class=""tfx"">"

		arrFN = Split(fields,",")
		xcols = Ubound(arrFN)

		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand    = CreateObject("ADODB.Command")

		objConnection.Provider = "ADsDSOObject"
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

					Select Case Ucase(fieldname)
						Case "NAME":
							fn = "Name"
							fv = strValue
							funx = CMWT_NameParse (cn)
						Case "DESCRIPTION":
							fn = "Description"
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
						Case "L":
							fn = "City"
							fv = strvalue
						Case "ST":
							fn = "State"
							fv = strvalue
						Case "FACSIMILETELEPHONENUMBER":
							fn = "FAX"
							fv = strvalue
						Case Else:
							fv = strValue
							fn = fieldname
					End Select

					Response.Write "<tr class=""tr1"">" & _
						"<td class=""td6 v10 bgGray w200"">" & fn & "</td>" & _
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
		
		Response.Write "<table class=""tfx"">"

		groups = Enum_Groups (UserID)
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
End Select

CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>