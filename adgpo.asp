<!-- #include file="_core.asp" -->
<%
'-----------------------------------------------------------------------------
' filename....... adgpo.asp
' lastupdate..... 12/10/2016
' description.... group policy object details
'-----------------------------------------------------------------------------
time1 = Timer

GpoName = CMWT_GET("id", "")
PSet    = CMWT_GET("set", "1")
CMWT_VALIDATE GpoName, "GPO Name was not provided"
PageTitle    = "GPO: " & GpoName
PageBackLink = "adgpos.asp"
PageBackName = "Group Policy Objects"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "1=Properties,2=Settings"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	'aduser.asp?uid=" & UserID & "&set=2
	mlink = "adgpo.asp?id=" & GpoName & "&set=" & mset(0)
	If KeySet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

strDomain   = Application("CMWT_DOMAINSUFFIX")
strForest   = Application("CMWT_DOMAINSUFFIX")

On Error Resume Next
Set objGPM = CreateObject("GPMgmt.GPM")
If err.Number <> 0 Then
	CMWT_ERROR "Group Policy Management tools must be installed on the CMWT host server"
End If


Response.Write "<table class=""tjx"">"

Select Case PSet
	Case "1"
	
		Set objGPMConstants = objGPM.GetConstants()
		Set objGPMDomain = objGPM.GetDomain(strDomain, "", objGPMConstants.UseAnyDC)
		Set objGPMSitesContainer = objGPM.GetSitesContainer(strForest, strDomain, "", objGPMConstants.UseAnyDC)
		Set objGPMSearchCriteria = objGPM.CreateSearchCriteria

		objGPMSearchCriteria.Add objGPMConstants.SearchPropertyGPODisplayName, objGPMConstants.SearchOpEquals, CStr(GpoName)
		Set objGPOList = objGPMDomain.SearchGPOs(objGPMSearchCriteria)

		xrows = objGPOList.Count

		fnlist = "DisplayName,DomainName,ModificationTime"

		Response.Write "<table class=""tfx""><tr>"
		For each fn in Split(fnlist, ",")
			Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
		Next
		for each objGPO in objGPOList
			dn = objGPO.DisplayName
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Name</td>" & _
				"<td class=""td6 v10"">" & GpoName & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">ID</td>" & _
				"<td class=""td6 v10"">" & objGPO.ID & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Domain Name</td>" & _
				"<td class=""td6 v10"">" & objGPO.DomainName & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Last Modified</td>" & _
				"<td class=""td6 v10"">" & objGPO.ModificationTime & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Computer DS Version</td>" & _
				"<td class=""td6 v10"">" & objGPO.ComputerDSVersionNumber & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">User DS Version</td>" & _
				"<td class=""td6 v10"">" & objGPO.UserDSVersionNumber & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Computer SysVol Version</td>" & _
				"<td class=""td6 v10"">" & objGPO.ComputerSysVolVersionNumber & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">User SysVol Version</td>" & _
				"<td class=""td6 v10"">" & objGPO.UserSysVolVersionNumber & "</td></tr>" & _
				"<tr class=""tr1"">" & _
				"<td class=""td6 v10 w200 bgGray"">Path</td>" & _
				"<td class=""td6 v10"">" & objGPO.Path & "</td></tr>"
		next

	Case "2"
		' display GPO settings
		strReportFile = "temp_gpo_" & Replace(GpoName," ", "_") & ".html"
		strReportPath = Application("CMWT_PhysicalPath") & "\scripts\" & strReportFile

		Set objGPMConstants = objGPM.GetConstants()
		Set objGPMDomain = objGPM.GetDomain(strDomain, "", objGPMConstants.UseAnyDC)
		Set objGPMSitesContainer = objGPM.GetSitesContainer(strForest, strDomain, "", objGPMConstants.UseAnyDC)
		Set objGPMSearchCriteria = objGPM.CreateSearchCriteria

		objGPMSearchCriteria.Add objGPMConstants.SearchPropertyGPODisplayName, objGPMConstants.SearchOpEquals, CStr(GpoName)
		Set objGPOList = objGPMDomain.SearchGPOs(objGPMSearchCriteria)

		Set objGPMResult = objGPOList.Item(1).GenerateReportToFile( objGPMConstants.ReportHTML, strReportPath )
		On Error Resume Next
		objGPMResult.OverallStatus()

		If objGPMResult.Status.Count > 0 Then
			for i = 1 to objGPMResult.Status.Count
				Response.Write "<tr><td class=""td6 v10"">" & objGPMResult.Status.Item(i).Message & "</td></tr>"
			next
		end if
		if err.Number = 0 then
			response.write "<tr class=""h100""><td class=""td6 v10 ctr"">" & _
				"<p>The GPO Settings have been exported to a report file and are ready for viewing.</p>" & _
				"<table><tr><td class=""m111 w250"" onClick=""javascript:window.open('scripts/" & strReportFile & "');"">" & _
				"Open Report File</td></tr></table></td></tr>"
		else
			response.write "<tr><td class=""td6 v10"">error=" & err.Number & ": " & err.Description & "</td></tr>"
		end if

	'Case "3"
		' display list of GPO links

		'Set objGPMSearchCriteria = objGPM.CreateSearchCriteria
		'objGPMSearchCriteria.Add objGPMConstants.SearchPropertySOMLinks, objGPMConstants.SearchOpContains, objGPOList.Item(1)
		'objSOMList = objGPMDomain.SearchSOMs(objGPMSearchCriteria)
		'Set objSiteLinkList = objGPMSitesContainer.SearchSites(objGPMSearchCriteria)
		'
		'If objSOMList.Count = 0 And objSiteLinkList.Count = 0 Then
		'	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No Site, Domain or OU links found for this GPO</td></tr>"
		'Else
		'	For each objSOM in objSOMList
		'		Select Case objSOM.Type
		'			Case objGPMConstants.SOMDomain
		'				strSOMType = "Domain"
		'			Case objGPMConstants.SOMOU
		'				strSOMType = "OU"
		'			Case Else
		'				strSOMType = "?"
		'		End Select
		'		Response.Write "<tr class=""tr1""><td class=""td6 v10"">" & objSOM.Name & "</td>" & _
		'			"<td class=""td6 v10"">" & strSOMType & "</td></tr>"
		'	Next
		'	For each objSiteLink in objSiteLinkList
		'		Response.Write "<tr class=""tr1""><td class=""td6 v10"">" & objSiteLink.Name & "</td>" & _
		'			"<td class=""td6 v10"">Site</td></tr>"
		'	Next
		'End If
End Select
Response.Write "</table>"
%>