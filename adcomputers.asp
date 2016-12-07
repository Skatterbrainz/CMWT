<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adComputers.asp
' lastupdate..... 11/30/2016
' description.... active directory computer accounts
'-----------------------------------------------------------------------------
time1 = Timer

oupath  = CMWT_GET("ou", Application("CMWT_DomainPath"))
objType = CMWT_GET("type", "user")
objPfx  = CMWT_GET("ch", "A")
QueryON = CMWT_GET("qq", "")

findFN = CMWT_GET("f", "")
findFV = CMWT_GET("v", "")

CMWT_VALIDATE oupath, "AD LDAP path was not specified"
If objPFX = "ALL" Then
	PageTitle = "Computers: ALL"
Else
	PageTitle = "Computers: Beginning with " & objPFX
End If
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
'CMWT_PageHeading PageTitle, ""

CMWT_CLICKBAR objPfx, "adcomputers.asp?ch="

Response.Write "<table class=""tfx"">"

Dim adConn, objComment, adRS, x, d
Dim retval : retval = ""
Dim fields, i, fieldname, strvalue, query

On Error Resume Next

fields = "logonCount,pwdLastSet,whenCreated,operatingSystemServicePack,operatingSystem,description,name"
query = "SELECT " & fields & " FROM 'LDAP://" & ouPath & "' " & _
	"WHERE objectCategory='computer'"
If objPFX <> "ALL" Then
	query = query & " AND name='" & objPFX & "*'"
End If
reflink = "device.asp?cn="

Response.Write "<tr>"
arrFN = Split(fields,",")
xcols = Ubound(arrFN)

For i = xcols to 0 Step -1
	Response.Write "<td class=""td6 v10 bgGray"">" & arrFN(i) & "</td>"
Next
Response.Write "</tr>"

Set adConn = CreateObject("ADODB.Connection")
Set adCmd  = CreateObject("ADODB.Command")

adConn.Provider = "ADsDSOObject"
adConn.Properties("ADSI Flag") = 1
adConn.Open "Active Directory Provider"

Set adCmd.ActiveConnection = adConn

adCmd.Properties("Page Size") = 1000
adCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
adCmd.CommandText = query

Set adRS = adCmd.Execute
adRS.MoveFirst
xrows = adRS.RecordCount

If xrows > 0 Then

	Do Until adRS.EOF

		Response.Write "<tr class=""tr1"">"

		For i = 0 to adRS.Fields.Count -1
			fieldname = adRS.Fields(i).Name
			strvalue  = adRS.Fields(i).Value

			Select Case Ucase(fieldname)
				Case "NAME":
					fv = strValue
					funx = CMWT_NameParse (cn)
					If Ucase(objType) = "GROUP" Then
						fv = "<a href=""" & reflink & NetSuffix & "\" & fv & """ title=""Details for " & fv & """>" & fv & "</a>"
					Else
						fv = "<a href=""" & reflink & fv & """ title=""Details for " & fv & """>" & fv & "</a>"
					End If
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
				'Case "WHENCREATED": 
				'	fv = LargeIntegerToDate(strvalue)
				Case "DISTINGUISHEDNAME":
					fv = strvalue
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
				Case "DESCRIPTION":
					If NotNullString(strValue) Then
						d = ""
						For each x in strValue
							d = d & x
						Next
						fv = d
					Else
						fv = ""
					End If
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
				Case Else:
					fv = strValue
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select

		Next
		Response.Write "</tr>"

		adRS.MoveNext
	Loop

	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & _
		xrows & " items were returned</td></tr>"
Else
	Response.Write "<tr class=""h100""><td class=""td6 v10 ctr"" colspan=""" & xcols+1 & """>" & _
		"No matching accounts were found</td></tr>"
End If
	
Response.Write "</table>"
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>
