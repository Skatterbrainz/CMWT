<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: appdplist.asp
' Date......: 12/04/2016
' Purpose...: application content distribution report
'****************************************************************
time1 = Timer

AppName = CMWT_GET("app","")
SortBy  = CMWT_GET("s", "DPName")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE AppName, "Application Name was not specified"

PageTitle    = "Distribution of " & AppName
PageBackLink = "dpservers.asp"
PageBackName = "Distribution Points"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx"">"
		
query = "SELECT DISTINCT DPName, " & _
	"Installed,Retrying,Failed,LastUpdated " & _
	"FROM dbo.vDPStatusPerDP " & _
	"WHERE SoftwareName='" & AppName & "' " & _
	"ORDER BY " & SortBy

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
	
	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			
			Select Case Ucase(fn)
				Case "DPNAME":
					fvv = CMWT_CN(fv)
					Response.Write "<td class=""td6 v10"">" & _
						"<a href=""dpapplist.asp?dp=" & fvv & """ title=""Show Applications"">" & fv & "</a></td>"
				Case "QTY":
					total = total + fv
					Response.Write "<td class=""td6 v10 right"">" & fv& "</td>"
				Case "INSTALLED":
					c_inst = c_inst + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case "RETRYING":
					c_retr = c_retr + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case "FAILED":
					c_fail = c_fail + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case "LASTUPDATED":
					Response.Write "<td class=""td6 v10 w200 right"">" & fv & "</td>"
				Case Else:
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select	
		Next
		rs.MoveNext
	Loop
	
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"">" & xrows & " rows returned</td>" & _
		"<td class=""td6 v10 w80 right bgGreen"">" & c_inst & "</td>" & _
		"<td class=""td6 v10 w80 right bgBlue"">" & c_retr & "</td>" & _
		"<td class=""td6 v10 w80 right bgOrange"">" & c_fail & "</td>" & _
		"<td class=""td6 v10 bgGray w200""></td>" & _
		"</tr>"
End If

Response.Write "</table>"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>
	
</body>
</html>