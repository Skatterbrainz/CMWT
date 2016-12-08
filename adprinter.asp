<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adprinter.asp
' lastupdate..... 11/30/2016
' description.... AD printer details
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer
pn = CMWT_GET("id", "")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE pn, "Printer name was not specified"

PageTitle = "Printer: " & pn
PageBackLink = "adprinters.asp"
PageBackName = "Shared Printers"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Response.Write "<table class=""tfx"">"

query = "SELECT whenCreated, printColor, printCollate, " & _
	"printDuplexSupported, printStaplingSupported, printKeepPrintedJobs," & _
	"printLanguage, printMediaReady, printPagesPerMinute, driverName, " & _
	"url, location, serverName, printerName, name " & _
	"FROM 'LDAP://" & Application("CMWT_DOMAINSUFFIX") & _
	"' WHERE objectClass='printQueue' AND name='" & pn & "'"

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand    = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = query
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute

If Not(objRecordSet.BOF AND objRecordSet.EOF) Then
	objRecordSet.MoveFirst
	On Error Resume Next
	Do Until objRecordSet.EOF
		For i = 0 to objRecordSet.Fields.Count - 1
			fn = objRecordSet.Fields(i).Name
			fv = objRecordSet.Fields(i).Value
			If VarType(fv) > 8 Then
				fv = Join(fv, "|")
			End If
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">" & fn & "</td>" & _
				"<td class=""td6 v10"">" & fv & "</td></tr>"
		Next
		objRecordSet.MoveNext
	Loop
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
		"No matching rows returned</td></tr>"
End If

Response.Write "</table>"

CMWT_Footer()
%>

</body>
</html>