<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... adgpos.asp
' lastupdate..... 12/10/2016
' description.... group policy objects
'-----------------------------------------------------------------------------
time1 = Timer
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","DisplayName")

PageTitle    = "Group Policy Objects"
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
strDomain   = Application("CMWT_DOMAINSUFFIX")

On Error Resume Next
Set objGPM = Server.CreateObject("GPMgmt.GPM")
If err.Number <> 0 Then
	CMWT_ERROR "Group Policy Management tools must be installed on the CMWT host server"
End If

Set objGPMConstants = objGPM.GetConstants()
Set objGPMDomain = objGPM.GetDomain(strDomain, "", objGPMConstants.UseAnyDC)
Set objGPMSearchCriteria = objGPM.CreateSearchCriteria
Set objGPOList = objGPMDomain.SearchGPOs(objGPMSearchCriteria)

xrows = objGPOList.Count

fnlist = "DisplayName,DomainName,ModificationTime"

Response.Write "<table class=""tfx""><tr>"
For each fn in Split(fnlist, ",")
	Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
Next
Response.Write "</tr>"
for each objGPO in objGPOList
	dn = objGPO.DisplayName
	response.write "<tr class=""tr1"">"
	Response.Write "<td class=""td6 v10""><a href=""adgpo.asp?id=" & dn & """ title=""View Details"">" & dn & "</a></td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.ID & "</td>"
	Response.Write "<td class=""td6 v10"">" & objGPO.DomainName & "</td>"
	Response.Write "<td class=""td6 v10"">" & objGPO.ModificationTime & "</td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.ComputerDSVersionNumber & "</td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.UserDSVersionNumber & "</td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.ComputerSysVolVersionNumber & "</td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.UserSysVolVersionNumber & "</td>"
	'Response.Write "<td class=""td6 v10"">" & objGPO.Path & "</td>"
	Response.Write "</tr>"
next
Response.Write "</table>"

'CMWT_SHOW_QUERY()
CMWT_FOOTER() 
Response.Write "</body></html>"
%>