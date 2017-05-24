<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... compstatus2.asp
' lastupdate..... 05/23/2017
' description.... component status summary - detailed view
'-----------------------------------------------------------------------------
time1 = Timer

FilterFN  = CMWT_GET("cn", "")
FilterFV  = CMWT_GET("fv", "")
LinkQual  = CMWT_GET("lq", "")
QueryOn   = CMWT_GET("qq", "")

PageTitle    = FilterFV
PageBackLink = "compstatus.asp"
PageBackName = "Component Status"

Select Case LinkQual
	Case "2147483648"
		cstype = "Warnings"
	Case "1073741824"
		cstype = "Errors"
End Select

PageTitle = PageTitle & " (" & cstype & ")"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Dim conn, cmd, rs

query = "SELECT " & _
		"com.SiteCode, " & _
		"com.MachineName, " & _
		"stat.MessageID, " & _
		"stat.Time," & _
		"stat.ProcessID," & _
		"com.ComponentName " & _
	"FROM " & _
		"v_StatusMessage stat " & _
			"JOIN v_ServerComponents com " & _
			"ON stat.SiteCode=com.SiteCode " & _
			"AND stat.MachineName=com.MachineName " & _
			"AND stat.Component=com.ComponentName " & _
	"WHERE " & _
		"Time > DATEADD(ss,-240-(24*3600),GETDATE()) " & _
		"AND " & _
		"com.ComponentName='" & FilterFV & "' " & _
		"AND " & _ 
		"Severity='-" & LinkQual & "'"

CMWT_DB_QUERY Application("DSN_CMDB"), query

if not (rs.BOF and rs.EOF) then 
	xrows = rs.RecordCount 
	xcols = rs.Fields.Count
	Response.Write "<table class=""tfx""><tr>"
	for i = 0 to xcols -1
		fn = rs.fields(i).name
		Select Case Ucase(fn)
			Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
				Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
			Case Else:
				Response.Write "<td class=""td6 v10 bgGray"">"
		End Select
		Response.Write CMWT_SORTLINK("compstatus2.asp?id=" & CompName, fn, SortBy) & "</td>"
	next
	Response.Write "</tr>"
	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			fv = "<a href=""compstatus2.asp?id=" & CompName & "&fn=" & fn & "&fv=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
			Response.Write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
		next
		rs.MoveNext
	Loop
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & _
		xrows & " rows returned</td></tr></table>"
end if

CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()
Response.Write "</body></html>"
%>
