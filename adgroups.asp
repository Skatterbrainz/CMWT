<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'****************************************************************
' Filename..: adgroups.asp
' Date......: 12/13/2016
' Purpose...: active directory groups list
'****************************************************************
time1 = Timer

objPfx  = CMWT_GET("ch", "A")
QueryOn = CMWT_GET("qq", "")

If CMWT_NotNullString(fxn) Then
	filtered = TRUE
End If

PageTitle    = "Security Groups"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "adgroups.asp?ch="
	
Response.Write "<table class=""tfx"">"

query = "SELECT ADsPath, Name FROM 'LDAP://" & Application("CMWT_DomainPath") & "' WHERE objectCategory='group'"

If objPFX <> "ALL" Then
	query = query & " AND name='" & objPFX & "*'"
End If

On Error Resume Next
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Properties("User ID")  = Application("CM_AD_TOOLUSER")
objConnection.Properties("Password") = Application("CM_AD_TOOLPASS")
objConnection.Properties("ADSI Flag") = 1
objConnection.Open "Active Directory Provider"

Set objCommand = CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCommand.CommandText = query

Set objRecordSet = objCommand.Execute
If Not(objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.MoveFirst

	Set rs = CreateObject("ADODB.RecordSet")
	rs.CursorLocation = adUseClient
	rs.Fields.Append "name", adVarChar, 255
	rs.Fields.Append "adspath", adVarChar, 255
	rs.Open
	Do Until objRecordSet.EOF
		gn = objRecordSet.Fields("Name").Value
		ap = objRecordSet.Fields("aDsPath").value
		rs.AddNew
		rs.Fields("name").value = gn
		rs.Fields("adspath").value = ap
		rs.Update
		objRecordSet.MoveNext
	Loop
	rs.Sort = "name"
	rs.MoveFirst

	xrows = objRecordSet.RecordCount
	
	Response.Write "<tr><td class=""td6 v10 bgGray"">Name</td>" & _
		"<td class=""td6 v10 bgGray"">Description</td></tr>"

	Do Until rs.EOF
		Set objGroup = GetObject(rs.Fields("ADsPath").Value)
		desc = objGroup.Description
		If Len(desc) > 120 Then
			desc = Left(desc, 120) & "..."
		End If
		gn = rs.Fields("Name").Value
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10"">" & _
			"<a href=""adgroup.asp?gn=" & gn & """>" & gn & "</a></td>" & _
			"<td class=""td6 v10"">" & desc & "</td></tr>"
		rs.MoveNext
	Loop
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""2"">" & _
		xrows & " rows returned</td></tr>"
	
	rs.Close
	Set rs = Nothing
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
		"No matching rows returned</td></tr>"
End If

objRecordSet.Close
objConnection.Close

Response.Write "</table>"
CMWT_Footer()
CMWT_SHOW_QUERY()
%>

</body>
</html>