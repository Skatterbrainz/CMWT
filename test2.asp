<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... test2.asp
' lastupdate..... 12/07/2016
' description.... ASP SQL connection test validation
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle    = Application("CMWT_SubTitle")
PageBackLink = ""
PageBackName = ""
SelfLink     = "test2.asp"

Dim conn
On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeOut = 5
conn.Open Application("DSN_CMDB")
If err.Number <> 0 Then
	testresult = "Failed!"
Else
	testresult = "Passed!"
End If
conn.Close
err.Clear

CMWT_NewPage "", "", ""

Response.Write "<span style=""font-size:30px;color:#00995c"">ConfigMgr Web Tools</span>"

Response.Write "<table class=""tfx"">" & _
	"<tr>" & _
		"<td class=""td6 v10"">" & _
			"<h1>Welcome to CMWT!</h1>" & _
			"<p class=""cMedBlue"">CMWT Site Testing Process</p>" & _
			"<table class=""tf800"">" & _
				"<tr>" & _
					"<td class=""td6a v10 w80 ctr bgGreen""><a href=""test.htm"">HTML</a></td>" & _
					"<td class=""td6a v10"">Passed!</td>" & _
				"</tr>" & _
				"<tr>" & _
					"<td class=""td6a v10 w80 ctr bgGreen""><a href=""test.asp"">ASP</a></td>" & _
					"<td class=""td6a v10"">Passed!</td>" & _
				"</tr>" & _
				"<tr>" & _
					"<td class=""td6a v10 w80 ctr bgGreen"">Database</td>" & _
					"<td class=""td6a v10"">" & testresult & "</td>" & _
				"</tr>" & _
			"</table>" & _
		"</td></tr></table>" & _
		"<br/><p class=""ctr"">" & _
		"<input type=""button"" name=""b1"" id=""b1"" class=""btx w140 h30"" value=""Open CMWT!"" title=""Go to CMWT Home Page"" onClick=""document.location.href='./'"" /></p>"

	CMWT_FOOTER()

Response.Write "</body></html>"
%>