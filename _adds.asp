<!-- begin-module: _adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... _adds.asp
' lastupdate..... 11/30/2016
' description.... active directory domain services module
'-----------------------------------------------------------------------------

Function CMWT_UAC (intVal)
	Select Case intVal
		Case 512:	 CMWT_UAC = "Enabled Account"
		Case 514:	 CMWT_UAC = "Disabled Account"
		Case 544:	 CMWT_UAC = "Enabled, Password Not Required"
		Case 546:	 CMWT_UAC = "Disabled, Password Not Required"
		Case 66048:	 CMWT_UAC = "Enabled, Password Doesnt Expire"
		Case 66050:	 CMWT_UAC = "Disabled, Password Doesnt Expire"
		Case 66080:	 CMWT_UAC = "Enabled, Password Doesnt Expire / Not Required"
		Case 66082:	 CMWT_UAC = "Disabled, Password Doesnt Expire / Not Required"
		Case 262656: CMWT_UAC = "Enabled, Smartcard Required"
		Case 262658: CMWT_UAC = "Disabled, Smartcard Required"
		Case 262688: CMWT_UAC = "Enabled, Smartcard Required, Password Not Required"
		Case 262690: CMWT_UAC = "Disabled, Smartcard Required, Password Not Required"
		Case 328192: CMWT_UAC = "Enabled, Smartcard Required, Password Doesnt Expire"
		Case 328194: CMWT_UAC = "Disabled, Smartcard Required, Password Doesnt Expire"
		Case 328224: CMWT_UAC = "Enabled, Smartcard Required, Password Doesnt Expire / Not Required"
		Case 328226: CMWT_UAC = "Disabled, Smartcard Required, Password Doesnt Expire / Not Required"
		Case Else: CMWT_UAC = "UNKNOWN"
	End Select
End Function

'----------------------------------------------------------------
' function-name: EnumerateOUs
' function-desc: 
'----------------------------------------------------------------

Function EnumerateOUs ()
	Dim conn, rs, ObjRootDSE, Leaf
	Dim StrSQL, StrDomName, ObjOU, Result : result = 0
	Set ObjRootDSE = GetObject("LDAP://RootDSE") 
	StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext")) 
	Set ObjRootDSE = Nothing 
	StrSQL = "SELECT Name, ADsPath FROM 'LDAP://" & StrDomName & "' WHERE ObjectCategory = 'OrganizationalUnit' AND Name <> 'Domain Controllers'" 
	Set conn = CreateObject("ADODB.Connection") 
	conn.Provider = "ADsDSOObject":    conn.Open "Active Directory Provider" 
	Set rs = CreateObject("ADODB.Recordset") 
	rs.Open StrSQL, conn 
	If Not rs.EOF Then
		result = rs.RecordCount
		rs.MoveLast
		rs.MoveFirst 
		While Not rs.EOF 
			Set ObjOU = GetObject(Trim(rs.Fields("ADsPath").Value)) 
			If StrComp(Right(Trim(ObjOU.Parent), Len(Trim(ObjOU.Parent)) - 7), StrDomName, VbTextCompare) = 0 Then 
				Leaf = Trim(rs.Fields("Name").Value)
				Response.Write "<option value=""""><strong>" & Leaf & "</strong></option>"
				GetChild(ObjOU) 
			End If         
			rs.MoveNext 
			Set ObjOU = Nothing 
		Wend 
	End If 
	rs.Close
	Set rs = Nothing 
	conn.Close
	Set conn = Nothing
	EnumerateOUs = result
End Function

'----------------------------------------------------------------
' function-name: GetChild
' function-desc: 
'----------------------------------------------------------------

Private Sub GetChild (ThisObject) 
	Dim ObjChild, StrThisParent, LeafName1, LeafPath1, LeafY
	For Each ObjChild In ThisObject 
		If StrComp(Trim(ObjChild.Class), "OrganizationalUnit", VbTextCompare) = 0 Then 
			LeafName1 = Right(Trim(ObjChild.Name), Len(Trim(ObjChild.Name)) - 3)
			LeafPath1 = ObjChild.ADsPath
			LeafY = OULink(LeafName1, LeafPath1)
			Response.Write "<option value=""" & LeafY & """ title=""View Contents..."">&nbsp;&nbsp;" & LeafName1 & "</option>"
			GetGrandChild (ObjChild.ADsPath) 
		End If         
	Next 
End Sub 

'----------------------------------------------------------------
' function-name: GetGrandChild
' function-desc: 
'----------------------------------------------------------------

Private Sub GetGrandChild (ThisADsPath) 
	Dim ObjGrand, ObjItem, subDN, LeafName2, LeafPath2, LeafX
	Set ObjGrand = GetObject(ThisADsPath) 
	For Each ObjItem In ObjGrand 
		subDN = objItem.ADsPath
		If StrComp(Trim(ObjItem.Class), "OrganizationalUnit", VbTextCompare) = 0 Then 
			LeafName2 = Right(Trim(ObjItem.Name), Len(Trim(ObjItem.Name)) - 3)
			LeafX = OULink(LeafName2, SubDN)
			Response.Write "<option value=""" & LeafX & """ title=""View Contents...."">&nbsp;&nbsp;&nbsp;&nbsp;" & LeafName2 & "</option>"
		End If 
		GetGrandChild Trim(ObjItem.ADsPath) 
	Next     
	Set ObjGrand = Nothing 
End Sub

'----------------------------------------------------------------
' function-name: OULink
' function-desc: 
'----------------------------------------------------------------

Function OULink (LabelName, ou)
	Select Case Ucase(LabelName)
		Case "WORKSTATIONS","SERVERS","DEVICES","POINTOFSALE","TRAININGCOMPUTERS","VIRTUAL DESKTOPS":
			OULink = "accounts.asp?type=computer&ou=" & Replace(ou, "LDAP://", "")
		Case "FSD","FDD","AWC","AZC":
			OULink = "accounts.asp?type=computer&ou=" & Replace(ou, "LDAP://", "")
		Case "GROUPS","SECURITY","APPLICATIONS","DISTRIBUTION":
			OULink = "accounts.asp?type=group&ou=" & Replace(ou, "LDAP://", "")
		Case "USERS","TERMINATED USERS","SERVICEACCOUNTS","ADMINS":
			OULink = "accounts.asp?type=user&ou=" & Replace(ou, "LDAP://", "")
		Case Else:
			OULink = "accounts.asp?type=user&ou=" & Replace(ou, "LDAP://", "")
	End Select
End Function

'----------------------------------------------------------------
' function-name: Enum_Groups
' function-desc: 
'----------------------------------------------------------------

Function Enum_Groups (strUserName)
	Dim result : result = ""
	Dim objRootDSE, strDomName, strSQL, objConn, objRS, GroupCollection, objGroup, objUser
	Set ObjRootDSE = GetObject("LDAP://RootDSE") 
	StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext")) 
	Set ObjRootDSE = Nothing
	StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & StrUserName & "'" 
	Set ObjConn = CreateObject("ADODB.Connection") 
	ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
	Set ObjRS = CreateObject("ADODB.Recordset") 
	ObjRS.Open StrSQL, ObjConn 
	If Not ObjRS.EOF Then 
		ObjRS.MoveLast:    ObjRS.MoveFirst 
		Set ObjUser = GetObject (Trim(ObjRS.Fields("ADsPath").Value)) 
		Set GroupCollection = ObjUser.Groups 
		For Each ObjGroup In GroupCollection 
			If result <> "" Then
				result = result & "," & objGroup.CN
			Else
				result = objGroup.CN
			End If
			CheckForNestedGroup ObjGroup 
		Next 
		Set ObjGroup = Nothing:    Set GroupCollection = Nothing:    Set ObjUser = Nothing 
	End If 
	ObjRS.Close:    Set ObjRS = Nothing 
	ObjConn.Close:    Set ObjConn = Nothing 
	Enum_Groups = result
End Function

'----------------------------------------------------------------
' function-name: CheckForNestedGroup
' function-desc: 
'----------------------------------------------------------------

Private Sub CheckForNestedGroup (ObjThisGroupNestingCheck) 
    On Error Resume Next 
    Dim AllMembersCollection, StrMember, StrADsPath, ObjThisIsNestedGroup 
    Dim result : result = ""
    AllMembersCollection = ObjThisGroupNestingCheck.GetEx("MemberOf") 
    For Each StrMember in AllMembersCollection 
        StrADsPath = "LDAP://" & StrMember 
        Set ObjThisIsNestedGroup = GetObject(StrADsPath) 
        If result <> "" Then
        	result = result & ",(" & Trim(ObjThisIsNestedGroup.CN) & ")"
        Else
        	result = "(" & Trim(ObjThisIsNestedGroup.CN) & ")"
        End If
        CheckForNestedGroup(ObjThisIsNestedGroup) 
    Next 
    Set ObjThisIsNestedGroup = Nothing:    Set StrMember = Nothing:    Set AllMembersCollection = Nothing 
End Sub

'----------------------------------------------------------------
' function-name: Get_ADsPath
' function-desc: 
'----------------------------------------------------------------

Function Get_ADsPath (strName, objType)
	Dim query, conn, cmd, rs, result : result = ""
	query = "SELECT distinguishedName FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
		"WHERE objectCategory='" & objType & "' AND name='" & strName & "'"
	Set conn = CreateObject("ADODB.Connection")
	Set cmd    = CreateObject("ADODB.Command")
	conn.Provider = "ADsDSOObject"
	conn.Properties("ADSI Flag") = 1
	conn.Open "Active Directory Provider"
	Set cmd.ActiveConnection = conn
	cmd.Properties("Page Size") = 1000
	cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	cmd.CommandText = query
	Set rs = cmd.Execute
	rs.MoveFirst
	xrows = rs.RecordCount
	If xrows > 0 Then
		Do Until rs.EOF
			result = rs.Fields("distinguishedName").value
			rs.MoveNext
		Loop
	End If
	rs.Close
	conn.Close
	Set rs = Nothing : Set cmd = Nothing : Set conn = Nothing
	Get_ADsPath = result
End Function

'----------------------------------------------------------------
' function-name: Get_DisplayName
' function-desc: 
'----------------------------------------------------------------

Function Get_DisplayName (UserID)
	Dim uid, query, conn, cmd, rs, result : result = ""
	uid = Replace(Lcase(UserID), Lcase(NetSuffix) & "\", "")
	query = "SELECT displayName FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
		"WHERE objectCategory='user' AND sAMAccountName='" & uid & "'"
	Set conn = CreateObject("ADODB.Connection")
	Set cmd = CreateObject("ADODB.Command")
	conn.Provider = "ADsDSOObject"
	conn.Properties("ADSI Flag") = 1
	conn.Open "Active Directory Provider"
	Set cmd.ActiveConnection = conn
	cmd.Properties("Page Size") = 1000
	cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	cmd.CommandText = query
	Set rs = cmd.Execute
	rs.MoveFirst
	xrows = rs.RecordCount
	If xrows > 0 Then
		Do Until rs.EOF
			result = rs.Fields("displayName").value
			rs.MoveNext
		Loop
	End If
	rs.Close
	conn.Close
	Set rs = Nothing : Set cmd = Nothing : Set conn = Nothing
	Get_DisplayName = result
End Function

'----------------------------------------------------------------
' function-name: Get_LogonName
' function-desc: 
'----------------------------------------------------------------

Function Get_LogonName (UserDN)
	Dim objUser, result : result = ""
	Set objUser = GetObject("LDAP://" & UserDN)
	result = objUser.sAMAccountName
	Get_LogonName = result
End Function

%>
<!-- end-module: _adds.asp -->
