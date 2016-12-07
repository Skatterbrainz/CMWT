<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adAccounts.asp
' Author....: David M. Stein
' Date......: 11/30/2016
'****************************************************************
time1 = Timer

oupath  = CMWT_GET("ou", "")
objType = CMWT_GET("type", "user")
findFN  = CMWT_GET("f", "")
findFV  = CMWT_GET("v", "")

'CMWT_VALIDATE oupath, "AD LDAP path was not specified"
'CMWT_VALIDATE_LIST objType, "user,computer,group,all"
PageTitle = "AD Accounts: OU Members"

Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN " & _
	" http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & _
	"<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">" & _
	"<head><meta charset=""utf-8"">" & _
	"<meta http-equiv=""Content-Language"" content=""en-us"" />" & _
	"<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"" />" & _
	"<meta http-equiv=""Cache-Control"" content=""cache"""" />" & _
	"<meta name=""distribution"" content=""Global"" />" & _
	"<meta name=""revisit-after"" content=""1 days"" />" & _
	"<meta name=""robots"" content=""follow, index, noodp, noydir"" />" & _
	"<meta name=""description"" content="""" />" & _
	"<meta name=""abstract"" content="""" />" & _
	"<meta name=""author"" content=""David M. Stein"" />" & _
	"<meta name=""copyright"" content=""David M. Stein"" />" & _
	"<meta name=""keywords"" content="""" />" & _
	"<title>CMWT: " & PageTitle & "</title>" & _
	"<link rel=""stylesheet"" type=""text/css"" href=""default.css"" />" & _
	"<link rel=""shortcut icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
	"<link rel=""icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
	"<script src=""_cmwt.js""></script>" & mr & _
	"</head>" & _
	"<body style=""width:100%"">"

Response.Write "<table class=""tfx"">"

Dim objConnection, objComment, objRecordSet, x, d
Dim retval : retval = ""
Dim fields, i, fieldname, strvalue, query

On Error Resume Next

Select Case Ucase(objType)
	Case "COMPUTER":
		fields = "logonCount,pwdLastSet,whenCreated,operatingSystemServicePack,operatingSystem,name"
		query = "SELECT " & fields & " FROM 'LDAP://" & ouPath & "' " & _
			"WHERE objectCategory='computer'"
		If findFN <> "" And findFV <> "" Then
			query = query & " AND " & findFN & "='" & findFV & "*'"
		End If
		datalink = "device.asp?cn="
	Case "USER":
		fields = "pwdLastSet,whenCreated,displayName,samaccountname"
		query = "SELECT " & fields & " FROM 'LDAP://" & ouPath & "' " & _
			"WHERE objectCategory='user'"
		If findFN <> "" And findFV <> "" Then
			query = query & " AND " & findFN & "='" & findFV & "*'"
		End If
		datalink = "aduser.asp?uid="
	Case "GROUP":
		fields = "whenCreated,description,name"
		query = "SELECT " & fields & " FROM 'LDAP://" & ouPath & "'"
		If findFN <> "" And findFV <> "" Then
			query = query & " WHERE " & findFN & "='" & findFV & "*'"
		End If
		datalink = "group.asp?gn="
	Case Else:
		fields = "description,name"
		query = "SELECT " & fields & " FROM 'LDAP://" & ouPath & "'"
		If findFN <> "" And findFV <> "" Then
			query = query & " WHERE " & findFN & "='" & findFV & "*'"
		End If
		datalink = ""
End Select

If CMWT_NotNullString(query) Then

	Response.Write "<tr>"
	arrFN = Split(fields,",")
	xcols = Ubound(arrFN)
	
	For i = xcols to 0 Step -1
		Response.Write "<td class=""td6 v10 bgGray"">" & arrFN(i) & "</td>"
	Next
	Response.Write "</tr>"

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
					Case "SAMACCOUNTNAME":
						fv = strValue
						fv = "<a href=""" & datalink & fv & """ title=""Details for " & fv & """ target=""_top"">" & fv & "</a>"
					Case "NAME":
						fv = strValue
						'funx = CMWT_NameParse (cn)
						If Ucase(objType) = "COMPUTER" Then
							fv = "<a href=""" & datalink & fv & """ title=""Details for " & fv & """ target=""_top"">" & fv & "</a>"
						End If
					'Case "WHENCREATED": 
					'	fv = LargeIntegerToDate(strvalue)
					Case "DISTINGUISHEDNAME":
						fv = strvalue
					Case "DESCRIPTION":
						If CMWT_NotNullString(strValue) Then
							d = ""
							For each x in strValue
								d = d & x
							Next
							fv = d
						Else
							fv = ""
						End If
					Case Else:
						fv = strValue
				End Select

				Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			Next
			Response.Write "</tr>"

			objRecordSet.MoveNext
		Loop

		Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & _
			xrows & " rows were returned</td></tr>"
	Else
		Response.Write "<tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"" colspan=""" & xcols+1 & """>No matching accounts found</td></tr>"
	End If
End If
	
Response.Write "</table>"
CMWT_FOOTER()
%>
</body>
</html>
