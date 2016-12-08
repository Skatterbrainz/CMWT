<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dpapplist.asp
' lastupdate..... 12/04/2016
' description.... applications assigned to a given DP server
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer

DPName  = CMWT_GET("dp","")
SortBy  = CMWT_GET("s", "SoftwareName")
QueryON = CMWT_GET("qq", "")

CMWT_VALIDATE DPName, "No DP server name was specified"

DPShortName  = Split(DPName,".")(0)
PageTitle    = "Applications on " & DPShortName
PageBackLink = "dpservers.asp"
PageBackName = "Distribution Points"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx"">"

query = "SELECT DISTINCT SoftwareName,LastUpdated, " & _
	"Installed,Retrying,Failed " & _
	"FROM dbo.vDPStatusPerDP " & _
	"WHERE (DPName='" & DPName & "') " & _
	"AND (LTRIM(SoftwareName) <> '') " & _
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
				Case "SOFTWARENAME":
					Response.Write "<td class=""td6 v10"">" & _
						"<a href=""appdplist.asp?app=" & Server.URLEncode(fv) & _
						""" title=""Show Distribution Status for " & fv & """>" & fv & "</a></td>"
				Case "LASTUPDATED":
					Response.Write "<td class=""td6 v10 w200"">" & fv & "</td>"
				Case "INSTALLED":
					c_inst = c_inst + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case "RETRYING":
					c_retr = c_retr + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case "FAILED":
					c_fail = c_fail + fv
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case Else:
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select	
		Next
		rs.MoveNext
	Loop

	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""2"">" & xrows & " rows returned</td>" & _
		"<td class=""td6 v10 right w80 bgGreen"">" & c_inst & "</td>" & _
		"<td class=""td6 v10 right w80 bgBlue"">" & c_retr & "</td>" & _
		"<td class=""td6 v10 right w80 bgOrange"">" & c_fail & "</td>" & _
		"</tr>"
End If

CMWT_DB_CLOSE()

Response.Write "</table>"

CMWT_SHOW_QUERY()
CMWT_Footer()
%>
	
</body>
</html>