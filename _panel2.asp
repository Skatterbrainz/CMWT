<%
'-----------------------------------------------------------------------------
' filename....... _panel2.asp
' lastupdate..... 03/01/2017
' description.... CMWT home page dashboard panel
'-----------------------------------------------------------------------------
' comma-delimited list of WMI service names to ignore in report
IgnoreServices = "MapsBroker,sppsvc"

Sub WMI_TGRID (hostname, columns, className, clause, caption, sortby, autolink)
	Dim cn, objWMIService, colItems, objItem, query, val, PropertyName, afx, afn, afl, rows, cols
	Response.Write "<h2 class=""tfx"">" & caption & "</h2>" & _
		"<table class=""t1x""><tr>"
	cols = Ubound(Split(columns,","))+1
	For each cn in Split(columns, ",")
		Response.Write "<td class=""td5a v10 bgBlue"">" & _
			"<a href=""" & Request.ServerVariables("PATH_INFO") & _
			"?s=" & cn & """ title=""Sort Column"">" & cn & "</a></td>"
	Next
	Response.Write "</tr>"
	query = "SELECT " & columns & " FROM " & className
	If clause <> "" Then
		query = query & " WHERE " & clause 
	End If
	For each svc in Split(IgnoreServices,",")
		query = query & " AND (Name <> '" & svc & "')"
	Next
	Set objWMIService = GetObject("winmgmts:\\" & hostname & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery(query,,48)
	If CMWT_NotNullString(autolink) Then
		afx = Split(autolink,"=")
		afn = afx(0)
		afl = afx(1)
	Else
		afn = ""
		afl = ""
	End If
	For Each objItem in colItems
		Response.Write "<tr class=""tr2"">"
		For each PropertyName in Split(columns, ",")
			val = objItem.Properties_.Item(PropertyName)
			If CMWT_NotNullString(afn) And Ucase(afn)=Ucase(PropertyName) Then
				val = "<a href=""" & afl & "=" & val & """ title=""Click for Details"">" & val & "</a>"
			End If
			Response.Write "<td class=""td5a v10"">" & val & "</td>"
		Next
		Response.Write "</tr>"
		rows = rows + 1
	Next
	If rows = 0 Then
		Response.Write "<tr class=""h50"">" & _
			"<td class=""td5a v10"" colspan=""5"">" & _
			"<img src=""images/cmwt_check.png"" border=""0"" height=""45"" />" & _
			"All Services Appear Good</td></tr>"
	Else
		Response.Write "<tr><td class=""td5a v10 bgDarkGray"" colspan=""" & cols & """>" & _
			rows & " stopped services found (click Service Name for details)</td></tr>"
	End If
	Response.Write "</table>"
End Sub

wmi_columns = "DisplayName,Name,StartMode,State,StartName"
wmi_class   = "Win32_Service"
WMI_TGRID ".", wmi_columns, wmi_class, "StartMode='Auto' AND State='Stopped'", "Services Health", "DisplayName", "Name=service.asp?sn="

%>
