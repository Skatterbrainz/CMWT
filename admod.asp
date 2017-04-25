<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... admod.asp
' lastupdate..... 04/24/2017
' description.... active directory object modification
'-----------------------------------------------------------------------------
time1 = Timer

acctID  = CMWT_GET("acct", "")
destID  = CMWT_GET("dest", "")
typeID  = CMWT_GET("type", "")
QueryOn = CMWT_GET("qq", "")

CMWT_VALIDATE acctID, "Account ID was not specified"
CMWT_VALIDATE destID, "Destination ID was not specified"

'Response.Write "<p>action: " & typeID & "</p>"
'Response.Write "<p>account ID: " & acctID & "</p>"
'Response.Write "<p>destination: " & destID & "</p>"

If Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
	Dim conn, cmd, rs
	CMWT_DB_OPEN Application("DSN_CMWT")
End If

Select Case typeID
	Case "adduser"
		err.Clear
		On Error Resume Next
		udn = CMWT_AD_GetADsPath(acctID, "user")
		if err.number <> 0 Then response.write "exception1: " & err.Description : response.end

		gdn = CMWT_AD_GetADsPath(destID, "group")
		if err.number <> 0 Then response.write "exception2: " & err.Description : response.end

		set openDS  = GetObject("LDAP:")
		if err.number <> 0 Then response.write "exception3: " & err.Description : response.end

		set objUser  = openDS.OpenDSObject("LDAP://" & udn, Application("CM_AD_TOOLUSER"), Application("CM_AD_TOOLPASS"), ADS_SECURE_AUTHENTICATION)
		if err.number <> 0 Then response.write "exception4: " & err.Description : response.end

		set objGroup = openDS.OpenDSObject("LDAP://" & gdn, Application("CM_AD_TOOLUSER"), Application("CM_AD_TOOLPASS"), ADS_SECURE_AUTHENTICATION)
		if err.number <> 0 Then response.write "exception5: " & err.Description : response.end

		objGroup.Add("LDAP://" & udn)
		if err.number <> 0 Then response.write "exception6: " & err.Description : response.end
		
		CMWT_LogEvent conn, "AD", "GROUP MEMBER ADD", acctID& " has been added to " & destID
		
		targetUrl = "aduser.asp?uid=" & acctID & "&set=2"

	Case "remuser"
		'
End Select

If Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
	CMWT_DB_CLOSE()
End If

If targetUrl <> "" Then
	Response.Redirect targetUrl
End If
%>
