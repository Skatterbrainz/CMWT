<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'****************************************************************
' Filename..: addc.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: active directory domain controllers
'****************************************************************
time1 = Timer

PageTitle = "Domain Controllers"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

ADDom  = CMWT_GET("d", "")
SortBy = CMWT_GET("s", "Full_Domain_Name0")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Set objConnection = Server.CreateObject("ADODB.Connection")
Set objCommand = Server.CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
 
objCommand.CommandText = _
	"SELECT name,distinguishedName FROM " & _
	"'LDAP://ou=Domain Controllers," & Application("CMWT_DomainPath") & "' WHERE objectClass='computer'" 

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

Set objRecordSet = objCommand.Execute

If Not (objRecordSet.BOF and objRecordSet.EOF) Then
	objRecordSet.MoveFirst
	xrows = objRecordSet.RecordCount
	
	Response.Write "<table class=""tfx""><tr>" & _
		"<td class=""td6 v10 bgGray"">Server</td>" & _
		"<td class=""td6 v10 bgGray"">Distinguished Name</td></tr>"

	Set rs = CreateObject("ADODB.RecordSet")
		 
	rs.CursorLocation = adUseClient
	rs.Fields.Append "name", adVarChar, 255
	rs.Fields.Append "distinguishedName", adVarChar, 255
	rs.Open

	Do Until objRecordSet.EOF
		dn = objRecordSet.Fields("distinguishedName").Value
		cn = Replace(CMWT_CN(dn),"CN=","")
		
		rs.AddNew
		rs.Fields("name").value = cn
		rs.Fields("distinguishedname").value = dn
		rs.Update

		objRecordSet.MoveNext
	Loop

	rs.Sort = "name"
	rs.MoveFirst

	Do Until rs.EOF
		dn = rs.Fields("distinguishedName").Value
		cn = rs.Fields("name").value
		cnx = "<a href=""device.asp?cn=" & cn & """>" & cn & "</a>"
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10"">" & cnx & "</td>" & _
			"<td class=""td6 v10"">" & dn & "</td></tr>"
		rs.MoveNext
	Loop

	rs.Close
	Set rs = Nothing
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""2"">" & _
		xrows & " rows returned</td></tr></table>"
Else
	Response.Write "<table class=""tfx""><tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
		"No matching rows returned</td></tr></table>"
End If

CMWT_Footer()
%>

</body>
</html>
