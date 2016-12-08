<!-- #include file=_core.asp -->
<!-- #include file=_adds.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adsites.asp
' lastupdate..... 11/30/2016
' description.... active directory sites
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle = "Sites"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

Sub CMWT_AD_SUBNETS (sn)
	Dim objRootDSE, strSiteRDN, strSitePath, objSite
	Dim arrSiteObjectBL, strSiteObjectBL
	On Error Resume Next
	strSiteRDN = sn
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strConfigurationNC = objRootDSE.Get("configurationNamingContext")
	if Err.Number <> 0 Then
		Response.Write "<li>Error1: " & err.Number & " / " & err.Description & "</li>"
	End If
	strSitePath = "LDAP://" & strSiteRDN & ",cn=Sites," & strConfigurationNC
	Set objSite = GetObject(strSitePath)
	if Err.Number <> 0 Then
		Response.Write "<li>Error2: " & err.Number & " / " & err.Description & "</li>"
	End If
	objSite.GetInfoEx Array("siteObjectBL"), 0
	arrSiteObjectBL = objSite.GetEx("siteObjectBL")
	Response.Write "<ul>"
	if Err.Number <> 0 Then
		Response.Write "<li>Error3: " & err.Number & " / " & err.Description & "</li>"
	End If
	For Each strSiteObjectBL In arrSiteObjectBL
		Response.Write "<li>" & Split(Split(strSiteObjectBL, ",")(0), "=")(1) & "</li>"
	Next
	Response.Write "</ul>"
End Sub

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
On Error Resume Next

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")

strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.Filter = Array("site")

Response.Write "<table class=""tfx"">" & _
	"<tr>" & _
	"<td class=""td6 v10 bgGray"">Site Name</td>" & _
	"<td class=""td6 v10 bgGray"">Servers</td>" & _
	"<td class=""td6 v10 bgGray"">Subnets</td>" & _
	"</tr>"

xrows = 0
For Each objSite In objSitesContainer
	sn = objSite.CN
	fv = "<a href=""cmadsite.asp?sn=" & sn & """ title=""Query Computers in site: " & sn & """>" & sn & "</a>"
	Response.Write "<tr class=""tr1"">"
	Response.Write "<td class=""td6 v10"">" & fv & "</td>"
	strSiteName = objSite.Name
	strServerPath = "LDAP://cn=Servers," & strSiteName & ",cn=Sites," & strConfigurationNC
	Set colServers = GetObject(strServerPath)

	Response.Write "<td class=""td6 v10""><ul>"
	For Each objServer In colServers
		ServerName = objServer.CN
		Response.Write "<li><a href=""device.asp?cn=" & ServerName & """>" & ServerName & "</a></li>"
	Next
	Response.Write "</ul></td>"
	Response.Write "<td class=""td6 v10"">"
	
	CMWT_AD_SUBNETS strSiteName
	
	Response.Write "</td></tr>"
	xrows = xrows + 1
Next

Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""3"">" & _
	xrows & " rows returned</td></tr></table>"

CMWT_Footer()
%>

</body>
</html>