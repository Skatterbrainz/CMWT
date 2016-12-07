<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: adattributes.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: active directory object schema attributes listing
'****************************************************************
Response.Expires = -1
time1 = Timer
StrClass = CMWT_GET("c", "user")

On Error Resume Next

Dim objTypeClass, objSchemaClass
Set objTypeClass = GetObject("LDAP://schema/" & StrClass)
Set objSchemaClass = GetObject(objTypeClass.Parent)

PageTitle = "AD Schema Attributes: " & Ucase(StrClass)
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

Private Sub GetAttributes(x,y,z)
	Dim strAttribute, strOut, objAttribute
	For Each strAttribute in x
		strOut = ""
		If z = True then
			strOut = strOut & "<td class=""td6 v10"">Yes</td>" & "<td class=""td6 v10"">" & strAttribute & "</td>"
		Else
			strOut = strOut & "<td class=""td6 v10"">No</td>" & "<td class=""td6 v10"">" & strAttribute & "</td>"
		End If
		Set objAttribute = y.GetObject("Property",  strAttribute)
		strOut = strOut & "<td class=""td6 v10"">" & objAttribute.Syntax & "</td>"
		If objAttribute.MultiValued Then
			strOut = strOut & "<td class=""td6 v10"">Multi</td>"
		Else
			strOut = strOut & "<td class=""td6 v10"">Single</td>"
		End If
		Response.Write "<tr class=""tr1"">" & strOut & "</tr>"
		strOut = Empty
		attCount = attCount + 1
	Next
	Set objAttribute = Nothing
	strAttribute = Empty
End Sub

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<div class=""tfx"">" & _
	"Select a Class to Explore: <select name=""x"" id=""x"" size=""1"" class=""w300 pad6"" " & _
	"onChange=""if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }"">" & _
	"<option value=""""></option>"
For each cx in Split("computer,contact,container,domainpolicy,group,organizationalunit,trusteddomain,user", ",")
	Response.Write "<option value=""adattributes.asp?c=" & cx & """>" & cx & "</option>"
Next
Response.Write "</select></div>"

attCount = 0
Response.Write "<table class=""tfx""><tr><td class=""pad6 v10"" colspan=""4""><h3>Mandatory Attributes</h3></td></tr>" & _
	"<tr>" & _
	"<td class=""td6 v10 bgGray"">Mandatory</td>" & _
	"<td class=""td6 v10 bgGray"">Name</td>" & _
	"<td class=""td6 v10 bgGray"">Syntax</td>" & _
	"<td class=""td6 v10 bgGray"">Single/Multi</td></tr>"
GetAttributes objTypeClass.MandatoryProperties, objSchemaClass, True
Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""4"">" & attCount & " mandatory attributes found</td></tr>"

attCount = 0
Response.Write "<tr><td class=""pad6 v10"" colspan=""4""><h3>Optional Attributes</h3></td></tr>" & _
	"<tr>" & _
	"<td class=""td6 v10 bgGray"">Mandatory</td>" & _
	"<td class=""td6 v10 bgGray"">Name</td>" & _
	"<td class=""td6 v10 bgGray"">Syntax</td>" & _
	"<td class=""td6 v10 bgGray"">Single/Multi</td></tr>"
GetAttributes objTypeClass.OptionalProperties, objSchemaClass, False
Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""4"">" & attCount & " optional attributes found</td></tr>"
Response.Write "</table>"
CMWT_FOOTER()
Response.Write "</body></html>"

Set objTypeClass = Nothing
Set objSchemaClass = Nothing

CMWT_Footer()
%>

</body>
</html>