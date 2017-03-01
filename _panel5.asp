<%
'-----------------------------------------------------------------------------
' filename....... _panel5.asp
' lastupdate..... 03/01/2017
' description.... CMWT home page dashboard panel
'-----------------------------------------------------------------------------
warning_num  = 0.85
critical_num = 0.95

wmi_class   = "Win32_LogicalDisk"
wmi_columns = "DeviceID,VolumeName,FileSystem,Size,Used,Free"
caption = "Disk Health"

Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3",,48) 

Response.Write "<h2 class=""tfx"">" & caption & "</h2>" & _
	"<table class=""t1x""><tr>"
cols = Ubound(Split(wmi_columns,","))+1
For each cn in Split(wmi_columns, ",")
	Response.Write "<td class=""td5a v10 bgBlue"">" & cn & "</td>"
Next
Response.Write "<td class=""td5a v10 bgBlue"">Status</td>"
Response.Write "</tr>"

rows = 0
For Each objItem in colItems
	x1 = objItem.DeviceID
	x2 = objItem.VolumeName
	x3 = objItem.FileSystem
	x4 = Round(objItem.Size / 1024.0 / 1024 / 1024, 2)
	x5 = Round(objItem.FreeSpace / 1024.0 / 1024 / 1024, 2)
	
	gbytes_used = x4 - x5
	pct_used = Round(1 - (x5 / x4), 2)
	If pct_used >= critical_num Then
		Response.Write "<tr class=""tr2"" onClick=""document.location.href='device.asp?cn=" & _
			Application("CMWT_SiteServer") & "&set=Logical Disks'"">" & _
			"<td class=""td5a v10"">" & x1 & "</td>" & _
			"<td class=""td5a v10"">" & x2 & "</td>" & _
			"<td class=""td5a v10"">" & x3 & "</td>" & _
			"<td class=""td5a v10"">" & x4 & " GB</td>" & _
			"<td class=""td5a v10"">" & gbytes_used & " GB</td>" & _
			"<td class=""td5a v10"">" & x5 & " GB</td>" & _
			"<td class=""td5a v10"">Critical</td></tr>"
		rows = rows + 1
	ElseIf pct_used >= warning_num Then
		Response.Write "<tr class=""tr2"" onClick=""document.location.href='device.asp?cn=" & _
			Application("CMWT_SiteServer") & "&set=Logical Disks'"">" & _
			"<td class=""td5a v10"">" & x1 & "</td>" & _
			"<td class=""td5a v10"">" & x2 & "</td>" & _
			"<td class=""td5a v10"">" & x3 & "</td>" & _
			"<td class=""td5a v10"">" & x4 & " GB</td>" & _
			"<td class=""td5a v10"">" & gbytes_used & " GB</td>" & _
			"<td class=""td5a v10"">" & x5 & " GB</td>" & _
			"<td class=""td5a v10"">Warning</td></tr>"
		rows = rows + 1
	End If
Next

If rows = 0 Then
	Response.Write "<tr class=""h50"">" & _
		"<td class=""td5a v10"" colspan=""7"">" & _
		"<img src=""images/cmwt_check.png"" border=""0"" height=""45"" />" & _
		"All Disks Appear in Good Condition</td></tr>"
Else
	Response.Write "<tr><td class=""td5a v10 bgDarkGray"" colspan=""7"">" & _
		rows & " issues found</td></tr>"
End If
Response.Write "</table>"
%>
