<%
'-----------------------------------------------------------------------------
' filename....... _panel1.asp
' lastupdate..... 02/28/2017
' description.... CMWT home page dashboard panel
'-----------------------------------------------------------------------------
Response.Write "<table class=""tfx""><tr><td><br/>"
q = "SELECT SiteCode,SiteName,Version," & _
	"ServerName,InstallDir FROM dbo.v_Site " & _
	"ORDER BY Type DESC, SiteCode"
CMWT_DB_QUERY Application("DSN_CMDB"), q
Response.Write "<table class=""t1x""><tr>"
For i = 0 to rs.Fields.Count - 1
	Response.Write "<td class=""td6a v10 bgBlue"">" & rs.Fields(i).Name & "</td>"
Next
Response.Write "<td class=""td6a v10 bgBlue"">Branch Name</td>"
Response.Write "</tr>"
Do Until rs.EOF
	Response.Write "<tr>"
	For i = 0 to rs.Fields.Count - 1
		Response.Write "<td class=""td6a v10 bgDarkGray"">" & rs.Fields(i).Value & "</td>"
	Next
	Response.Write "<td class=""td6a v10 bgDarkGray"">" & CMWT_CM_BuildName(rs.Fields("Version").value) & "</td>"
	Response.Write "</tr>"
	rs.MoveNext
Loop
CMWT_DB_CLOSE()
Response.Write "</table></td></tr></table>"
%>
