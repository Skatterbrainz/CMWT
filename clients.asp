<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... clients.asp
' lastupdate..... 02/22/2017
' description.... computers with or without clients installed
'-----------------------------------------------------------------------------
time1 = Timer

cm = CMWT_GET("c", "")
QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s", "ComputerName")

Select Case cm
	Case "1": PageTitle = "Devices: With Client"
	Case "0": PageTitle = "Devices: No Client"
	Case Else: PageTitle = "Devices: Discovered"
End Select

PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Select Case cm
	Case "1":
		query = "SELECT DISTINCT " & _
			"Name0 AS ComputerName, ResourceID, " & _
			"AD_Site_Name0 AS ADSiteName, " & _
			"Full_Domain_Name0 AS DomainName, " & _
			"Client_Version0 AS ClientVersion, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System " & _
			"WHERE Client0 = 1"
	Case "0":
		query = "SELECT DISTINCT " & _
			"Name0 AS ComputerName, ResourceID, " & _
			"AD_Site_Name0 AS ADSiteName, " & _
			"Full_Domain_Name0 AS DomainName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System " & _
			"WHERE (Client0 IS NULL) OR (Client0 <> 1)"
	Case Else:
		query = "SELECT DISTINCT " & _
			"Name0 AS ComputerName, ResourceID, " & _
			"AD_Site_Name0 AS ADSiteName, " & _
			"Full_Domain_Name0 AS DomainName, " & _
			"Client_Version0 AS ClientVersion, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System"
End Select
query = query & " ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "clients.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
Response.Write "</body></html>"
%>
