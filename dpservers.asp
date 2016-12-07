<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: dpservers.asp
' Date......: 11/30/2016
' Purpose...: distribution point servers
'****************************************************************
Response.Expires = -1

PageTitle = "Distribution Points"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "ServerName")
QueryON = CMWT_GET("qq", "")

time1 = Timer

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx"">"

query = "SELECT REPLACE(DPName, '." & Application("CMWT_DOMAINSUFFIX") & "','') AS ServerName, DPName as FullName, " & _
	"Installed, Failed, Retry, Installed+Failed+Retry AS ITotal " & _
	"FROM (" & _
	"SELECT DISTINCT DPName, SUM(Installed) AS Installed, SUM(Failed) AS Failed, SUM(Retrying) AS Retry " & _
	"FROM dbo.vDPStatusPerDP " & _
	"GROUP BY DPName ) AS T1 " & _
	"ORDER BY ServerName"

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query

If Not(rs.BOF And rs.EOF) Then
	total = 0
	found = True
	xrows = rs.RecordCount
	xcols = rs.Fields.Count
	
	Response.Write "<tr>"
	For i = 0 to xcols-1
		fn = rs.Fields(i).Name
		Select Case Ucase(fn)
			Case "QTY":
				Response.Write "<td class=""td6 v10 w80 bgGray ctr"">" & fn & "</td>"
			Case Else:
				Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
		End Select
		
	Next
	Response.Write "</tr>"
	
	c_inst = 0
	c_retr = 0
	c_fail = 0
	c_totx = 0
	
	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			
			Select Case Ucase(fn)
				Case "SERVERNAME":
					Response.Write "<td class=""td6 v10"">" & _
						"<a href=""device.asp?cn=" & CMWT_CN(fv) & """ title=""Details for " & fv & """>" & fv & "</a></td>"
				Case "FULLNAME":
					Response.Write "<td class=""td6 v10"">" & _
						"<a href=""dpapplist.asp?dp=" & CMWT_CN(fv) & """ title=""Show Distribution Status for " & fv & """>" & fv & "</a></td>"
				Case "QTY","APPS":
					total = total + fv
					Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
				Case "INSTALLED":
					c_inst = c_inst + fv
					If fv > 0 Then
						Response.Write "<td class=""td6 v10 right bgGreen"">" & fv& "</td>"
					Else
						Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
					End If
				Case "RETRY":
					c_retr = c_retr + fv
					Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
				Case "FAILED":
					c_fail = c_fail + fv
					If fv > 0 Then
						Response.Write "<td class=""td6 v10 right bgLightRed"">" & fv& "</td>"
					Else
						Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
					End If
				Case "ITOTAL":
					c_totx = c_totx + fv
					Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
				Case Else:
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select	
		Next
		rs.MoveNext
	Loop
	
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""2"">" & _
		xrows & " rows returned</td>" & _
		"<td class=""td6 v10 bgGray right"">" & c_inst & "</td>" & _
		"<td class=""td6 v10 bgGray right"">" & c_retr & "</td>" & _
		"<td class=""td6 v10 bgGray right"">" & c_fail & "</td>" & _
		"<td class=""td6 v10 bgGray right"">" & c_totx & "</td>" & _
		"</tr>"
End If

Response.Write "</table>"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>
	
</body>
</html>