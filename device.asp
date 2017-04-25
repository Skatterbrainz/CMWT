<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... device.asp
' lastupdate..... 04/24/2017
' description.... device information report
'-----------------------------------------------------------------------------
time1 = Timer

cn = CMWT_GET("cn", "")
lx = CMWT_GET("logs", "")
tx = CMWT_GET("temp", "")
up = CMWT_GET("uprof", "")
QueryOn = CMWT_GET("qq", "")
pset    = CMWT_GET("set", "GENERAL")
SortBy  = CMWT_GET("s", "Name0")

CMWT_VALIDATE cn, "No device name was provided"

if InStr(cn,".") > 0 Then
	cnx = Split(cn,".")
	cn = cnx(0)
End If

PageTitle = cn
PageBackLink = "devices.asp"
PageBackName = "Devices"

If lx = "true" Then
	Show_LogFiles = True
End If
If tx = "true" Then
	Show_TempFiles = True
End If
If up = "true" Then
	Show_Profiles = True
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
Response.Write "<table class=""tfx"">" & _
	"<tr><td class=""td6"">" & _
	"<h2>" & cn & " : " & CMWT_MENUGROUP(pset) & "</h2>" & _
	"</td><td>"
CMWT_MENULIST pset, "device.asp?cn=" & cn
Response.Write "</td></tr></table>"

Dim conn, cmd, rs

Select Case Ucase(pset)

	Case "GENERAL":
					
		query = "SELECT TOP 1 " & _
				"Name0 AS [ComputerName], " & _
				"ResourceID, " & _
				"AD_Site_Name0 AS [ADSiteName], " & _
				"Manufacturer0 AS [Manufacturer], " & _
				"Model0 AS [ModelName], " & _
				"SystemType0 AS [SystemType], " & _
				"Client0 AS [Client], " & _
				"Operating_System_Name_and0 AS [WindowsType], " & _
				"Caption0 AS [WindowsVersion], " & _
				"CSDVersion0 AS [ServicePack], " & _
				"Domain0 AS [DomainName], " & _
				"PrimaryOwnerName0 AS [PrimaryOwner], " & _
				"TotalPhysicalMemory0 AS [Memory], " & _
				"ChassisTypes0 AS [ChassisType], " & _
				"SerialNumber0 AS [SerialNumber], " & _
				"UserName0 AS UserName, " & _
				"CurrentTimeZone0 AS [TimeZone] " & _
			"FROM (" & q_devices & ") AS T1 " & _
			"WHERE T1.Name0='" & cn & "'"

		CMWT_DB_QUERY Application("DSN_CMDB"), query

		Response.Write "<table class=""tfx"">"
		
		If Not(rs.BOF And rs.EOF) Then
			xcols = rs.Fields.Count
			xrows = rs.RecordCount

			Do Until rs.EOF
				For i = 0 to xcols-1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					Select Case Ucase(fn)
						Case "RESOURCEID":
							ResourceID = fv
						Case "ADSITENAME":
							fv = "<a href=""cmadsite.asp?sn=" & fv & """ title=""Devices in Site " & fv & """>" & fv & "</a>"
						Case "MODELNAME","MODEL0":
							fv = "<a href=""model.asp?m=" & fv & """ title=""Devices by Model " & fv & """>" & fv & "</a>"
'						Case "DOMAINNAME":
'							fv = "<a href=""domlist.asp?dn=" & fv & """ title=""Devices in " & fv & """>" & fv & "</a>"
						Case "MEMORY":
							fv = CMWT_KB2GB(fv) & " GB"
						Case "WINDOWSVERSION":
							If InStr(fv, "Server") > 0 Then
								is_server = True
							Else
								is_server = False
							End If
							fv = "<a href=""os.asp?on=" & fv & """ title=""Devices running " & fv & """>" & fv & "</a>"
						Case "CHASSISTYPE":
							fv = Get_ChassisType(fv, "")
							'fv = CMWT_CM_CHASSISTYPE(fv)
						Case "CLIENT":
							If fv = 1 Then
								fv = "Installed"
							Else
								fv = ""
							End If
					End Select

					Response.Write "<tr class=""tr1"">" & _
						"<td class=""td6 v10 bgGray w150"">" & fn & "</td>" & _
						"<td class=""td6 v10"">" & fv & "</td></tr>"
				Next
				If is_server Then
					Response.Write "<tr class=""tr1"">" & _
						"<td class=""td6 v10 bgGray w150"">OS Type</td>" & _
						"<td class=""td6 v10""><a href=""servers.asp"">Server</a></td>" & _
						"</tr>"
				End If
				ipa = CMWT_IP_BY_HOSTNAME (conn, cn)
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10 bgGray w150"">IP Address</td>" & _
					"<td class=""td6 v10"">" & ipa & "</td>" & _
					"</tr>"
				rs.MoveNext
			Loop

		Else
			Response.Write "<tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching record found</td></tr>"
		End If

		Response.Write "</table>"
		
		CMWT_DB_CLOSE()
		
	Case "SOFTWARE APPLICATIONS":
					
		ResourceID = CMWT_CM_RESOURCEID(cn)
		
		If ResourceID <> "" Then
		
			query = "SELECT DISTINCT " & _
				"DisplayName0 AS ProductName, Publisher0 AS Publisher " & _
				"FROM dbo.v_GS_ADD_REMOVE_PROGRAMS " & _
				"WHERE (ResourceID='" & ResourceID & "') " & _
				"AND (DisplayName0 IS NOT NULL) " & _
				"AND (LTRIM(DisplayName0)<> '') " & _
				"ORDER BY DisplayName0"

			CMWT_DEVICE_TABLE query
			
		Else
			Response.Write "<table class=""tfx"">" & _
				"<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
				"Resource ID was not found</td></tr></table>"
		End If
	
	Case "AGENT":
		
		query = "SELECT DISTINCT b.AgentName, " & _
			"b.AgentSite, b.AgentTime " & _
			"FROM dbo.v_R_System a INNER JOIN dbo.v_AgentDiscoveries b " & _
			"ON a.ResourceID=b.ResourceId " & _
			"WHERE Name0='" & cn & "' " & _
			"ORDER BY AgentName"
		
		CMWT_DEVICE_TABLE query
		
	Case "AUTOSTART":
		
		query = "SELECT DISTINCT " & _ 
			"dbo.v_GS_AUTOSTART_SOFTWARE.Product0 AS ProductName, " & _
			"dbo.v_GS_AUTOSTART_SOFTWARE.ProductVersion0 AS [Version], " & _
			"dbo.v_GS_AUTOSTART_SOFTWARE.Publisher0 AS Publisher, " & _
			"dbo.v_GS_AUTOSTART_SOFTWARE.StartupType0 AS StartType, " & _
			"dbo.v_GS_AUTOSTART_SOFTWARE.StartupValue0 AS StartValue " & _
			"FROM dbo.v_GS_AUTOSTART_SOFTWARE INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_AUTOSTART_SOFTWARE.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') " & _
			"ORDER BY ProductName"
			
		CMWT_DEVICE_TABLE query
	
	Case "DEPLOYMENTS":
		
		query = "SELECT DISTINCT dbo.v_FullCollectionMembership.CollectionID, " & _
			"dbo.vAppDeploymentStatus.AssignmentID, dbo.vCI_CIAssignments.AssignmentName, " & _
			"dbo.vCI_CIAssignments.StartTime " & _
			"FROM dbo.v_R_System INNER JOIN " & _
			"dbo.v_FullCollectionMembership ON " & _
			"dbo.v_R_System.ResourceID = dbo.v_FullCollectionMembership.ResourceID " & _
			"INNER JOIN dbo.vAppDeploymentStatus ON " & _
			"dbo.v_FullCollectionMembership.CollectionID = dbo.vAppDeploymentStatus.CollectionID " & _
			"INNER JOIN dbo.vCI_CIAssignments ON " & _
			"dbo.vAppDeploymentStatus.AssignmentID = dbo.vCI_CIAssignments.AssignmentID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') " & _
			"ORDER BY AssignmentName"
		
		CMWT_DEVICE_TABLE query
	
	Case "DUPLICATE FILES":
	
		query = "SELECT DISTINCT " & _
				"dbo.v_GS_SoftwareFile.FileName, " & _
				"COUNT(dbo.v_GS_SoftwareFile.FileName) AS Copies, " & _
				"SUM(dbo.v_GS_SoftwareFile.FileSize) AS TotalSize " & _
			"FROM " & _
				"dbo.v_R_System INNER JOIN " & _
				"dbo.v_GS_SoftwareFile ON dbo.v_R_System.ResourceID = dbo.v_GS_SoftwareFile.ResourceID " & _
			"WHERE " & _
				"(dbo.v_R_System.Name0 = '" & cn & "') " & _
			"GROUP BY FileName " & _
			"ORDER BY Copies DESC"

		CMWT_DEVICE_TABLE query
			
	Case "FEATURES":
	
		query = "SELECT DISTINCT dbo.v_GS_SERVER_FEATURE.Name0 AS FeatureName, dbo.v_R_System.Name0 AS Computername " & _
			"FROM dbo.v_GS_SERVER_FEATURE INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_SERVER_FEATURE.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') " & _
			"ORDER BY FeatureName"

		CMWT_DEVICE_TABLE query
		
	Case "LOGICAL DISKS":
				
		query = "SELECT DISTINCT Description0, " & _
			"DeviceID0 AS Drive, " & _
			"DriveType0 AS DriveType, " & _
			"FileSystem0 AS FileSystem, " & _
			"FreeSpace0 AS FreeSpace, " & _
			"MediaType0 AS MediaType, " & _
			"Size0 AS Capacity, " & _
			"VolumeName0 AS Label, " & _
			"VolumeSerialNumber0 AS VolumeSN " & _
			"FROM dbo.v_GS_LOGICAL_DISK " & _
			"WHERE SystemName0 = '" & cn & "'"

		CMWT_DEVICE_TABLE query
	
	Case "AUDIO DEVICES":
		
		query = "SELECT DISTINCT dbo.v_GS_SOUND_DEVICE.Name0 AS DeviceName, " & _
			"dbo.v_GS_SOUND_DEVICE.Manufacturer0 AS Manufacturer " & _
			"FROM dbo.v_GS_SOUND_DEVICE INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_SOUND_DEVICE.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE dbo.v_R_System.Name0='" & cn & "' " & _
			"ORDER BY DeviceName"
		
		CMWT_DEVICE_TABLE query
		
	Case "USER PROFILES":
					
		query = "SELECT DISTINCT " & _
			"dbo.v_R_User.User_Name0 AS UserID, " & _
			"dbo.v_GS_USER_PROFILE.LocalPath0 AS LocalPath " & _
			"FROM dbo.v_R_User INNER JOIN " & _
			"dbo.v_GS_USER_PROFILE ON dbo.v_R_User.SID0 = dbo.v_GS_USER_PROFILE.SID0 LEFT OUTER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_USER_PROFILE.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE dbo.v_R_System.Name0='" & cn & "' " & _
			"ORDER BY dbo.v_R_User.User_Name0"
		
		CMWT_DEVICE_TABLE query
	
	Case "SHARES":
		
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_SHARE.Name0 AS ShareName, dbo.v_GS_SHARE.Path0 AS PathName, " & _
			"dbo.v_GS_SHARE.Description0 AS Description " &_
			"FROM dbo.v_GS_SHARE INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_SHARE.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE dbo.v_R_System.Name0='" & cn & "'" & _
			"ORDER BY dbo.v_GS_SHARE.Name0"
		
		CMWT_DEVICE_TABLE query
					 
	Case "COLLECTIONS":
				
		query = "SELECT dbo.v_ClientCollectionMembers.CollectionID, " & _
			"dbo.v_Collection.Name, dbo.v_Collection.Comment, " & _
			"dbo.v_CollectionRuleQuery.CollectionID AS QZ " & _
			"FROM dbo.v_ClientCollectionMembers INNER JOIN " & _
			"dbo.v_Collection ON dbo.v_ClientCollectionMembers.CollectionID = " & _
			"dbo.v_Collection.CollectionID LEFT OUTER JOIN " & _
			"dbo.v_CollectionRuleQuery ON dbo.v_ClientCollectionMembers.CollectionID = " & _
			"dbo.v_CollectionRuleQuery.CollectionID " & _
			"WHERE (dbo.v_ClientCollectionMembers.Name = '" & cn & "') " & _
			"ORDER BY Name"

		CMWT_DB_QUERY Application("DSN_CMDB"), query
		
		Response.Write "<table class=""tfx"">"
		
		If Not(rs.BOF And rs.EOF) Then
			found = True
			xrows = rs.RecordCount
			xcols = rs.Fields.Count

			Response.Write "<tr>" & _
				"<td class=""td6 v10 w30 bgGray""> </td>" & _
				"<td class=""td6 v10 w80 bgGray"">ID</td>" & _
				"<td class=""td6 v10 bgGray"">Name</td>" & _
				"<td class=""td6 v10 bgGray"">Comment</td>" & _
				"</tr>"

			Do Until rs.EOF
				x1 = rs.Fields("CollectionID").value
				x2 = rs.Fields("Name").value
				x3 = rs.Fields("QZ").value
				x4 = rs.Fields("Comment").value
				
				Response.Write "<tr class=""tr1"">"
				Response.Write "<td class=""td6 v10 ctr"">"
				'If NotNullString(x3) Then
				'	CMWT_IMGLINK2 False, "icon_del2", "icon_del1", "icon_del3", "cmcx.asp?cid=" & x1 & "&cn=" & cn & "&mx=rem&z=device", "Query-Rule Collection"
				'Else
				'	CMWT_IMGLINK2 CMWT_ADMIN(), "icon_del2", "icon_del1", "icon_del3", "cmcx.asp?cid=" & x1 & "&cn=" & cn & "&mx=rem&z=device", "Remove"
				'End If
				Response.Write "</td>"
				
				Response.Write "<td class=""td6 v10"">" & x1 & "</td>"
				Response.Write "<td class=""td6 v10"">" & _
					"<a href=""collection.asp?id=" & x1 & """ title=""Collection Details"">" & x2 & "</a></td>"
				Response.Write "<td class=""td6 v10"">" & x4 & "</td>"
				Response.Write "</tr>"

				rs.MoveNext
			Loop
			Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
				xrows & " items were found</td></tr>"
		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No rows were found</td></tr>"
		End If
		
		Response.Write "<table>"
		
		If CMModify = True And CMWT_ADMIN() Then
			Response.Write "<form name=""form3"" id=""form3"" method=""post"" action=""cmcx.asp"">" & _
				"<input type=""hidden"" name=""cn"" id=""cn"" value=""" & cn & """ />" & _
				"<input type=""hidden"" name=""mx"" id=""mx"" value=""ADD"" />" & _
				"<table class=""tfx""><tr><td class=""td6 v10"">" & _
				"<select name=""cid"" id=""cid"" size=""1"" class=""w400 pad6"">" & _
					"<option value=""""></option>"

			clist = ""
			CMWT_CM_ListCollections conn, "", 2, clist

			Response.Write "</select> " & _
				"<input type=""submit"" name=""bx1"" id=""bx1"" class=""w140 h32 btx"" value=""Add"" />" & _
				"</td></tr></table>" & _
				"</form>"
		End If
		
		CMWT_DB_CLOSE()

	Case "ACTIVE DIRECTORY": 
	
		Dim objConnection, objComment, objRecordSet, x, d
		Dim retval : retval = ""
		Dim fields, i, fieldname, strvalue, query

		On Error Resume Next

		fields = "distinguishedName,logonCount,pwdLastSet,whenCreated,operatingSystemServicePack,operatingSystem,description,name"

		query = "SELECT " & fields & " FROM 'LDAP://" & Application("CMWT_DomainPath") & "' " & _
			"WHERE objectCategory='computer' AND name='" & cn & "'"

		Set objConnection = CreateObject("ADODB.Connection")
		Set objCommand    = CreateObject("ADODB.Command")

		Response.Write "<table class=""tfx"">"
		
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("ADSI Flag") = 1
		objConnection.Open "Active Directory Provider"

		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = 1000
		objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
		objCommand.CommandText = query

		Set objRecordSet = objCommand.Execute

		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			For i = 0 to objRecordSet.Fields.Count -1
				fieldname = objRecordSet.Fields(i).Name
				strvalue  = objRecordSet.Fields(i).Value
				Select Case Ucase(fieldname)
					Case "WHENCREATED": 
						strvalue = LargeIntegerToDate(strvalue)
					Case "DISTINGUISHEDNAME":
						strDN = strvalue
						strOU = Replace(Lcase(strDN), "cn=" & Lcase(cn) & ",", "")
						strValue = "<a href=""adaccounts.asp?type=computer&ou=" & strOU & """>" & strDN & "</a>"
					Case "DESCRIPTION":
						If NotNullString(strValue) Then
							d = ""
							For each x in strValue
								d = d & x
							Next
							strValue = d
						Else
							strValue = ""
						End If
				End Select
				
				Response.Write "<tr class=""tr1""><td class=""td6 v10 w180 bgGray"">" & fieldname & "</td>" & _
					"<td class=""td6 v10"">" & strvalue & "</td></tr>"
			Next
			objRecordSet.MoveNext
		Loop
		Response.Write "<tr class=""tr1""><td class=""td6 v10 w180 bgGray"">ADSI Path</td>" & _
			"<td class=""td6 v10"">" & CMWT_DN_ADSI(strDN) & "</td></tr>"
	
		Response.Write "</table>"
	
	Case "TOOLS":
		
		If CMWT_BROWSER_TYPE() = "IE" Then
						
			Response.Write "<table class=""tfx"">" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b11"" id=""b11"" class=""w140 h32 btx"" value=""Explore"" onClick=""javascript:explorer('" & cn & "');"" />" & _
					" Explore C: Drive on remote computer" & _
					"</td><td class=""td6a v10 right bgDarkGray"">file://" & cn & "/c$/" & _
					"</td></tr>" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b12"" id=""b12"" class=""w140 h32 btx"" value=""CCM Logs"" onClick=""javascript:ccmlogs('" & cn & "');"" />" & _
					" View CCM Client Logs on remote computer" & _
					"</td><td class=""td6a v10 right bgDarkGray"">\\" & cn & "\c$\windows\ccm\logs\" & _
					"</td></tr>" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b13"" id=""b13"" class=""w140 h32 btx"" value=""RDP"" onClick=""javascript:rdp('" & cn & "');"" />" & _
					" Connect to remote computer via Remote Desktop" & _
					"</td><td class=""td6a v10 right bgDarkGray"">mstsc -v " & cn & "" & _
					"</td></tr>" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b14"" id=""b14"" class=""w140 h32 btx"" value=""Manage"" onClick=""javascript:manage('" & cn & "');"" />" & _
					" Open Computer Management console for remote computer" & _
					"</td><td class=""td6a v10 right bgDarkGray"">compmgmt.msc -a /computer=" & cn & "" & _
					"</td></tr>" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b15"" id=""b15"" class=""w140 h32 btx"" value=""Command"" onClick=""javascript:winrs('" & cn & "');"" />" & _
					" Interactive Command Console for remote computer" & _
					"</td><td class=""td6a v10 right bgDarkGray"">winrs -r:" & cn & " cmd.exe" & _
					"</td></tr>" & _
				"<tr><td class=""td6a v10 bgDarkGray"">" & _
					"<input type=""button"" name=""b16"" id=""b16"" class=""w140 h32 btx"" value=""Help?"" onClick=""javascript:cmwthelp();"" />" & _
					" Help" & _
					"</td><td class=""td6a v10 right bgDarkGray"">" & _
					"</td></tr></table>"

		Else

			Response.Write "<table class=""tfx"">" & _
				"<tr class=""h100 tr1"">" & _
				"<td class=""td6a v10 ctr"">" & _
				"Remote Tools are only available for IE browsers." & _
				"<br/><br/>Tips:<br/>" & _
				"RDP = mstsc -v " & cn & "<br/>" & _
				"WinRS = winrs -r:" & cn & " cmd<br/>" & _
				"Computer Management = compmgmt.msc -a /computer=" & cn & _
				"</td></tr></table>"
						
		End If
	
	Case "LOCAL PRINTERS":
		
		query = "SELECT Name0 AS Name, Description0 AS Comment, DeviceID0 AS DeviceID, " & _
			"DriverName0 AS Driver, ShareName0 AS ShareName, Status0 AS Status " & _
			"FROM (" & q_printers & ") AS T1 WHERE T1.ComputerName='" & cn & "' " & _
			"ORDER BY " & SortBY

		CMWT_DEVICE_TABLE query
		
	Case "NETWORK ADAPTERS":
	
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DefaultIPGateway0 AS GateWay, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DHCPEnabled0 AS DHCP_Enabled, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DHCPServer0 AS DHCP_Server, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DNSDomain0 AS Domain, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DNSDomainSuffixSearchOrder0 AS DNS_Suffixes, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.DNSServerSearchOrder0 AS DNS_Servers, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.IPAddress0 AS IPAddress, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.MACAddress0 AS MAC, " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.Index0 AS [Index], " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.IPSubnet0 AS Subnet " & _
			"FROM dbo.v_R_System INNER JOIN " & _
			"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION ON dbo.v_R_System.ResourceID = dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') AND (dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.IPEnabled0 = 1)"
		
		CMWT_DEVICE_TABLE query
		
	Case "NOTES":
	
		query = "SELECT NoteID, Comment, DateCreated, CreatedBy " & _
			"FROM dbo.Notes " & _
			"WHERE (AttachedTo = '" & cn & "') AND (AttachClass = 'COMPUTER') " & _
			"ORDER BY NoteID DESC"
		
		Response.Write "<table class=""tfx"">"
		
		CMWT_DB_QUERY Application("DSN_CMWT"), query

		If Not(rs.BOF And rs.EOF) Then
			found = True
			xrows = rs.RecordCount
			xcols = rs.Fields.Count

			Response.Write "<tr>"
			For i = 0 to xcols-1
				Response.Write "<td class=""td6 v10 bgGray"">" & rs.Fields(i).Name & "</td>"
			Next
			Response.Write "</tr>"

			Do Until rs.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to xcols-1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					Select Case Ucase(fn)
						Case "NOTEID":
							fv = CMWT_IMG_LINK (TRUE, "icon_del2", "icon_del1", "icon_del3", "confirm.asp?id=" & fv & "&tn=notes&pk=noteid&t=device.asp|cn=" & cn & "^set=10", "Remove") & " " & _
								CMWT_IMG_LINK (TRUE, "icon_edit2", "icon_edit1", "icon_edit2", "noteedit.asp?id=" & fv, "Edit")
							Response.Write "<td class=""td6 v10 w50"">" & fv & "</td>"
						Case Else:
							Response.Write "<td class=""td6 v10"">" & fv & "</td>"
					End Select
					
				Next
				Response.Write "</tr>"
				rs.MoveNext
			Loop
			Response.Write "<tr>" & _
				"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
				xrows & " rows returned</td></tr>"
		Else
			Response.Write "<tr class=""h100 tr1"">" & _
				"<td class=""td6 v10 ctr"">No custom notes have been attached to this item</td></tr>"
		End If
		
		Response.Write "</table>"
		
		If CMWT_ADMIN() Then
			Response.Write "<br/><table class=""tfx""><tr><td class=""v10"">" & _
				"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" " & _
				"value=""New Note"" onClick=""document.location.href='noteadd.asp?id=" & cn & "&t=computer'"" " & _
				"title=""New Note"" /></td></tr></table>"
		End If
	
	Case "BIOS":
		
		query = "SELECT TOP 1 " & _
			"dbo.v_GS_PC_BIOS.Name0 AS BIOSName, dbo.v_GS_PC_BIOS.Version0 AS Version, " & _
			"dbo.v_GS_PC_BIOS.Manufacturer0 AS Manufacturer, dbo.v_GS_PC_BIOS.ReleaseDate0 AS ReleaseDate, " & _
			"dbo.v_GS_PC_BIOS.SerialNumber0 AS SerialNum, dbo.v_GS_PC_BIOS.SMBIOSBIOSVersion0 AS BIOSVersion " & _
			"FROM dbo.v_GS_PC_BIOS INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_PC_BIOS.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE dbo.v_R_System.Name0='" & cn & "'"
		
		CMWT_DEVICE_TABLE query
	
	Case "VIDEO":
	
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_VIDEO_CONTROLLER.Name0 AS Controller, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.AdapterRAM0 AS VRAM, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.CurrentBitsPerPixel0 AS BPXL, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.DeviceID0 AS DeviceID, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.DriverDate0 AS DriverDate, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.DriverVersion0 AS DriverVersion, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.InstalledDisplayDrivers0 AS Drivers, " & _
			"dbo.v_GS_VIDEO_CONTROLLER.VideoModeDescription0 AS VideoMode " & _
			"FROM dbo.v_GS_VIDEO_CONTROLLER INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_VIDEO_CONTROLLER.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') AND (dbo.v_GS_VIDEO_CONTROLLER.VideoProcessor0 IS NOT NULL) " & _
			"ORDER BY Controller"
		
		CMWT_DEVICE_TABLE query
	
	Case "MEMORY":
		
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_PHYSICAL_MEMORY.Name0 AS Name, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.Manufacturer0 AS Manufacturer, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.Model0 AS Model, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.MemoryType0 AS MemType, dbo.v_GS_PHYSICAL_MEMORY.PartNumber0 AS PartNum, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.Capacity0 AS Capacity, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.BankLabel0 AS Bank, dbo.v_GS_PHYSICAL_MEMORY.DataWidth0 AS DataWidth, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.DeviceLocator0 AS Locator, " & _
			"dbo.v_GS_PHYSICAL_MEMORY.SerialNumber0 AS SerialNum, dbo.v_GS_PHYSICAL_MEMORY.Speed0 AS Speed " & _
			"FROM dbo.v_R_System INNER JOIN " & _
			"dbo.v_GS_PHYSICAL_MEMORY ON dbo.v_R_System.ResourceID = dbo.v_GS_PHYSICAL_MEMORY.ResourceID " & _
			"WHERE dbo.v_R_System.Name0='" & cn & "' " & _
			"ORDER BY dbo.v_GS_PHYSICAL_MEMORY.Name0"
		
		CMWT_DEVICE_TABLE query
	
	Case "SOFTWARE FILES":
	
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_SoftwareFile.FileName, dbo.v_GS_SoftwareFile.FileDescription, dbo.v_GS_SoftwareFile.FilePath, " & _
			"dbo.v_GS_SoftwareFile.FileSize, dbo.v_GS_SoftwareFile.ModifiedDate, " & _
			"dbo.v_GS_SoftwareFile.FileModifiedDate, dbo.v_GS_SoftwareFile.FileVersion " & _
			"FROM dbo.v_R_System INNER JOIN " & _
			"dbo.v_GS_SoftwareFile ON dbo.v_R_System.ResourceID = dbo.v_GS_SoftwareFile.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') " & _
			"ORDER BY dbo.v_GS_SoftwareFile.FileName"
		
		CMWT_DEVICE_TABLE query

	Case "SOFTWARE UPDATES":
	
		query = "SELECT DISTINCT " & _
			"dbo.v_GS_QUICK_FIX_ENGINEERING.HotFixID0 AS HotFixID, dbo.v_GS_QUICK_FIX_ENGINEERING.Description0 AS Category, " & _
			"dbo.v_GS_QUICK_FIX_ENGINEERING.Caption0 AS Article, dbo.v_GS_QUICK_FIX_ENGINEERING.InstalledBy0 AS InstalledBy, " & _
			"dbo.v_GS_QUICK_FIX_ENGINEERING.InstalledOn0 AS InstallDate " & _
			"FROM dbo.v_GS_QUICK_FIX_ENGINEERING INNER JOIN " & _
			"dbo.v_R_System ON dbo.v_GS_QUICK_FIX_ENGINEERING.ResourceID = dbo.v_R_System.ResourceID " & _
			"WHERE (dbo.v_R_System.Name0 = '" & cn & "') " & _
			"ORDER BY HotFixID"
		
		CMWT_DEVICE_TABLE query
		
	Case "PHYSICAL DISKS":
	
		query = "SELECT DISTINCT " & _
			"Model0 AS Model,Index0 AS DriveID,InterfaceType0 AS Interface,MediaType0 AS Media, " & _
			"Partitions0 AS Partitions,SCSIBus0 AS Bus,SCSILogicalUnit0 AS LUN,SCSIPort0 AS Port, " & _
			"SCSITargetId0 AS Target,Size0 AS Capacity " & _
			"FROM dbo.v_GS_DISK " & _
			"WHERE SystemName0='" & cn & "' " & _
			"ORDER BY Model0"
		
		CMWT_DEVICE_TABLE query
	
	Case "PROCESSORS":
	
		query = "SELECT DISTINCT " & _
			"Name0 AS Name,DataWidth0 AS DataWidth,DeviceID0 AS DeviceID,IsHyperthreadCapable0 AS HyperThread, " & _
			"IsTrustedExecutionCapable0 AS TPM,IsVitualizationCapable0 AS VC,MaxClockSpeed0 AS MaxClock, " & _
			"NormSpeed0 AS NormSpeed,NumberOfCores0 AS Cores,NumberOfLogicalProcessors0 AS LogicalCPUs," & _
			"Revision0 AS Revision,SocketDesignation0 AS Socket " & _
			"FROM dbo.v_GS_PROCESSOR " & _
			"WHERE SystemName0='" & cn & "' " & _
			"ORDER BY Name0"
		
		CMWT_DEVICE_TABLE query
	
	Case Else:
		
		Response.Write "<table class=""tfx"">" & _
			"<tr class=""h100 tr1""><td class=""td6 v10 ctr"">" & _
			"Invalid Property Code</td></tr></table>"
		
End Select

CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()

Response.Write "</body></html>"
%>
