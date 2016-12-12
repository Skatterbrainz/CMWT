<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... service.asp
' lastupdate..... 12/09/2016
' description.... display windows service status
'-----------------------------------------------------------------------------
time1 = Timer
SvcName = CMWT_GET("sn", "")
SortBy  = CMWT_GET("s", "Name")
QueryON = CMWT_GET("qq", "")

CMWT_VALIDATE SvcName, "Service Name was not provided"

PageTitle    = SvcName
PageBackLink = "services.asp"
PageBackName = "Services"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
strComputer = "." 
query = "SELECT * FROM Win32_Service WHERE Name='" & SvcName & "'"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery(query,,48)

Response.Write "<table class=""tfx"">"
For Each objItem in colItems
	If objItem.StartMode = "Auto" And objItem.State = "Stopped" Then
		xxx = "&nbsp;<span style=""color:red"">*** WARNING! ***</span>"
	Else
		xxx = ""
	End If
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Display Name</td>" & _
		"<td class=""td6"">" & objItem.DisplayName & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Service Name</td>" & _
		"<td class=""td6"">" & objItem.Name & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Path Name</td>" & _
		"<td class=""td6"">" & objItem.PathName & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Start Mode</td>" & _
		"<td class=""td6"">" & objItem.StartMode & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">State</td>" & _
		"<td class=""td6"">" & objItem.State & xxx & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Status</td>" & _
		"<td class=""td6"">" & objItem.Status & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Start Name</td>" & _
		"<td class=""td6"">" & objItem.StartName & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Description</td>" & _
		"<td class=""td6"">" & objItem.Description & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Service Type</td>" & _
		"<td class=""td6"">" & objItem.ServiceType & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Accept Pause</td>" & _
		"<td class=""td6"">" & objItem.AcceptPause & "</td></tr>" & _
		"<tr class=""tr1"">" & _
		"<td class=""w180 td6 bgGray"">Accept Stop</td>" & _
		"<td class=""td6"">" & objItem.AcceptStop & "</td></tr>" 
Next
Response.Write "</table>"

CMWT_SHOW_QUERY() 
CMWT_Footer()
Response.Write "</body></html>"
%>
