<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adprinters.asp
' lastupdate..... 11/30/2016
' description.... AD printers report
'-----------------------------------------------------------------------------
Response.Expires = -1

objPfx  = CMWT_GET("ch", "A")
SortBy  = CMWT_GET("s", "printerName")
QueryOn = CMWT_GET("qq", "")

If CMWT_NotNullString(fxn) Then
	filtered = TRUE
End If

PageTitle = "Shared Printers"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

time1 = Timer

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "adprinters.asp?ch="

Response.Write "<table class=""tfx"">"

If objPFX <> "ALL" Then
	query = "SELECT location, shortServerName, name, printerName " & _
		"FROM 'LDAP://" & Application("CMWT_DOMAINPATH") & "' WHERE objectClass='printQueue' AND name='" & objPFX & "*'"
Else
	query = "SELECT location, shortServerName, name, printerName " & _
		"FROM 'LDAP://" & Application("CMWT_DOMAINPATH") & "' WHERE objectClass='printQueue'"
End If
query = query & " ORDER BY " & SortBy

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = query
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute

If Not(objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.MoveFirst
	xrows = objRecordSet.RecordCount
	xcols = objRecordSet.Fields.Count
	Response.Write "<tr>"
	For i = 0 to xcols - 1
		fn = objRecordSet.Fields(i).name
		Response.Write "<td class=""td6 v10 bgGray""><a href=""adprinters.asp?ch=" & objPFX & "&s=" & fn & """>" & fn & "</a></td>"
	Next
	Response.Write "</tr>"

	On Error Resume Next
	Do Until objRecordSet.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = objRecordSet.Fields(i).name
			fv = objRecordSet.Fields(i).value
			Select Case Ucase(fn)
				Case "NAME":
					fv = "<a href=""adprinter.asp?id=" & fv & """ title=""Printer Details"">" & fv & "</a>"
			End Select
			Response.Write "<td class=""td6 v10"">" & fv & "</td>"
		Next
		Response.Write "</tr>"
		objRecordSet.MoveNext
	Loop
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
		xrows & " rows returned</td></tr>"
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
		"No matching rows returned</td></tr>"
End If

Response.Write "</table>"
CMWT_Footer()
%>

</body>
</html>