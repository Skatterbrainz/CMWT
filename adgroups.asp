<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adgroups.asp
' Date......: 11/30/2016
' Purpose...: active directory groups list
'****************************************************************
time1 = Timer

objPfx  = CMWT_GET("ch", "A")
QueryOn = CMWT_GET("qq", "")

If CMWT_NotNullString(fxn) Then
	filtered = TRUE
End If

PageTitle = "Security Groups"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="./_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "adgroups.asp?ch="
	
Response.Write "<table class=""tfx"">"

If objPFX <> "ALL" Then
	query = "SELECT ADsPath, Name FROM 'LDAP://" & Application("CMWT_DomainPath") & "' WHERE objectCategory='group' AND name='" & objPFX & "*'"
Else
	query = "SELECT ADsPath, Name FROM 'LDAP://" & Application("CMWT_DomainPath") & "' WHERE objectCategory='group'"
End If

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCommand.CommandText = query

Set objRecordSet = objCommand.Execute
If Not(objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.MoveFirst
	xrows = objRecordSet.RecordCount
	
	Response.Write "<tr><td class=""td6 v10 bgGray"">Name</td>" & _
		"<td class=""td6 v10 bgGray"">Description</td></tr>"

	Do Until objRecordSet.EOF
		Set objGroup = GetObject(objRecordSet.Fields("ADsPath").Value)
		desc = objGroup.Description
		If Len(desc) > 120 Then
			desc = Left(desc, 120) & "..."
		End If
		gn = objRecordSet.Fields("Name").Value
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10"">" & _
			"<a href=""adgroup.asp?gn=" & gn & """>" & gn & "</a></td>" & _
			"<td class=""td6 v10"">" & desc & "</td></tr>"
		objRecordSet.MoveNext
	Loop
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""2"">" & _
		xrows & " rows returned</td></tr>"
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
		"No matching rows returned</td></tr>"
End If

objRecordSet.Close
objConnection.Close

Response.Write "</table>"
CMWT_Footer()
%>

</body>
</html>