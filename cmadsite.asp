<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmadsite.asp
' lastupdate..... 12/11/2016
' description.... computers assigned to specified AD site
'-----------------------------------------------------------------------------
time1 = Timer

ADSite  = CMWT_GET("sn", "")
Client  = CMWT_GET("c", "")
SortBy  = CMWT_GET("s", "ComputerName")
QueryOn = CMWT_GET("qq", "")

CMWT_VALIDATE ADSite, "Site Code was not specified"

If ADSite = "ALL" Then
	pageTitle = "All Computers"
ElseIf ADSite = "UNKNOWN" Then
	pageTitle = "Unassigned Computers"
Else
	pageTitle = "Computers by AD Site: " & ADSite
End If
PageBackLink = "adtools.asp"
PageBackName = "Active Directory"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Response.Write "<table class=""tfx"">"

If ADSite = "ALL" Then
	query = "SELECT DISTINCT Name0 AS ComputerName, " & _
		"Client0 AS Client, " & _
		"Caption0 AS WindowsType, " & _
		"Manufacturer0 AS Manufacturer, " & _
		"Model0 AS ModelName, " & _
		"TotalPhysicalMemory0 AS Memory, " & _
		"Domain0 AS DomainName, " & _
		"SystemType0 AS CPUType " & _
		"FROM (" & q_devices & ") AS T1"
	If Client = "1" Then
		query = query & " WHERE T1.Client0=1"
	End If

ElseIf ADSite = "UNKNOWN" Then
	QUERY = "SELECT DISTINCT " & _
		"dbo.v_R_System.Name0 AS ComputerName, " & _
		"dbo.v_R_System.AD_Site_Name0 AS ADSiteName, " & _
		"dbo.v_R_System.Client0 AS Client, " & _
		"dbo.v_GS_COMPUTER_SYSTEM.Domain0 AS DomainName, " & _
		"dbo.v_GS_COMPUTER_SYSTEM.Manufacturer0 AS Manufacturer, " & _
		"dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model, " & _
		"dbo.v_GS_X86_PC_MEMORY.TotalPhysicalMemory0 AS Memory, " & _
		"dbo.v_GS_OPERATING_SYSTEM.Caption0 AS QindowsType " & _
		"FROM dbo.v_R_System LEFT OUTER JOIN " & _
		"dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION ON " & _
		"dbo.v_R_System.ResourceID = dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.ResourceID LEFT OUTER JOIN " & _
		"dbo.v_GS_COMPUTER_SYSTEM ON dbo.v_R_System.Name0 = dbo.v_GS_COMPUTER_SYSTEM.Name0 LEFT OUTER JOIN " & _
		"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID LEFT OUTER JOIN " & _
		"dbo.v_GS_X86_PC_MEMORY ON dbo.v_R_System.ResourceID = dbo.v_GS_X86_PC_MEMORY.ResourceID " & _
		"WHERE (dbo.v_R_System.AD_Site_Name0 IS NULL) AND (dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.IPAddress0 IS NOT NULL)"
	If Client = "1" Then
		query = query & " AND dbo.v_R_System.Client0=1"
	End If

ElseIf Client = "1" Then
	query = "SELECT DISTINCT Name0 AS ComputerName, " & _
		"Client0 AS Client, " & _
		"Caption0 AS WindowsType, " & _
		"Manufacturer0 AS Manufacturer, " & _
		"Model0 AS ModelName, " & _
		"TotalPhysicalMemory0 AS Memory, " & _
		"Domain0 AS DomainName, " & _
		"SystemType0 AS CPUType " & _
		"FROM (" & q_devices & ") AS T1 " & _
		"WHERE T1.AD_Site_Name0='" & ADSite & "' " & _
		"AND T1.Client0=1"
	filtered = True

Else
	query = "SELECT DISTINCT Name0 AS ComputerName, " & _
		"Client0 AS Client, " & _
		"Caption0 AS WindowsType, " & _
		"Manufacturer0 AS Manufacturer, " & _
		"Model0 AS ModelName, " & _
		"TotalPhysicalMemory0 AS Memory, " & _
		"Domain0 AS DomainName, " & _
		"SystemType0 AS CPUType " & _
		"FROM (" & q_devices & ") AS T1 " & _
		"WHERE T1.AD_Site_Name0='" & ADSite & "'"
End If
query = query & " ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_OPEN Application("DSN_CMDB")
CMWT_DB_QUERY Application("DSN_CMDB"), query

If Not(rs.BOF And rs.EOF) Then
	found = True
	xrows = rs.RecordCount
	xcols = rs.Fields.Count
	
	Response.Write "<tr>"
	For i = 0 to xcols-1
		Response.Write "<td class=""td6 v10 bgGray"">" & rs.Fields(i).Name & "</td>"
	Next
	Response.Write "</tr>"

	invcount = 0
	
	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			
			Select Case Ucase(fn)
				Case "NAME0","COMPUTERNAME":
					fv = "<a href=""device.asp?cn=" & fv & """ title=""Details for " & fv & """>" & fv & "</a>"
				Case "MODEL0","MODELNAME","MODEL":
					If CMWT_NotNullString(fv) Then
						fv = "<a href=""model.asp?m=" & fv & """ title=""Computers by Model"">" & fv & "</a>"
						invcount = invcount + 1
					Else
						fv = ""
					End If
				Case "CLIENT","CLIENT0":
					If fv = 1 Then
						fv = "<a href=""cmadsite.asp?sn=" & ADSite & "&c=1"">Yes</a>"
					Else
						fv = ""
					End If
				Case "DOMAINNAME":
					fv = "<a href=""domlist.asp?dn=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
				Case "MEMORY":
					If fv > 0 Then
						fv = CMWT_KB2GB(fv) & " GB"
					Else
						fv = ""
					End If
			End Select
			
			Response.Write "<td class=""td6 v10"">" & fv & "</td>"
		Next
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	
	pct = FormatPercent( invCount / xrows, 2)

	If filtered = True Then
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
			xrows & " rows returned - (filtered list) " & _
			"<a href=""cmadsite.asp?sn=" & ADSite & """ title=""Show all in " & ADSite & """>Show All</a>" & _
			"</td></tr>"
	Else
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
			xrows & " items returned - " & invcount & " inventoried (" & pct & ")</td></tr>"
	End If
Else
	Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No computers were found in this AD site</td></tr>"
End If

Response.Write "</table>"

CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()
Response.Write "</body></html>"
%>
