<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... compstatus1.asp
' lastupdate..... 12/29/2016
' description.... component status summary
'-----------------------------------------------------------------------------
time1 = Timer

FilterFN  = CMWT_GET("fn", "")
FilterFV  = CMWT_GET("fv", "")
QueryOn   = CMWT_GET("qq", "")

PageTitle    = "Component Status"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Sub CMWT_DB_IntTableGrid (rs, Caption, LinkField, LinkQualifier)
	Dim xrows, xcols, fn, fv, i 
	if not (rs.BOF and rs.EOF) then 
		xrows = rs.RecordCount 
		xcols = rs.Fields.Count
		Response.Write "<h2 class=""tfx"">" & Caption & "</h2>"
		Response.Write "<table class=""tfx""><tr>"
		for i = 0 to xcols -1
			fn = rs.fields(i).name
			Select Case Ucase(fn)
				Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
					Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case Else:
					Response.Write "<td class=""td6 v10 bgGray"">"
			End Select
			Response.Write CMWT_SORTLINK("compstatus.asp", fn, SortBy) & "</td>"
		next
		Response.Write "</tr>"
		Do Until rs.EOF
			Response.Write "<tr class=""tr1"">"
			For i = 0 to xcols-1
				fn = rs.Fields(i).Name
				fv = rs.Fields(i).Value
				If Ucase(LinkField) = Ucase(fn) Then
					fv = "<a href=""compstatus2.asp?fn=" & LinkField & "&fv=" & fv & "&lq=" & LinkQualifier & _
						""" title=""Show Details"">" & fv & "</a>"
				End If
				Response.Write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
			next
			rs.MoveNext
		Loop
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & _
			xrows & " rows returned</td></tr></table>"
	end if
End Sub

Dim conn, cmd, rs

query = "SELECT " & _
		"T1.SiteCode, " & _
		"T1.ComponentName, " & _
		"T1.MachineName, " & _
		"SUM(CASE WHEN T1.Category='ERROR' THEN 1 ELSE 0 END) AS Errors, " & _
		"SUM(CASE WHEN T1.Category='WARNING' THEN 1 ELSE 0 END) AS Warnings, " & _
		"SUM(CASE WHEN T1.Category='INFO' THEN 1 ELSE 0 END) AS Info " & _
	"FROM ( " & _
	"SELECT DISTINCT " & _
		"com.SiteCode, " & _
		"com.ComponentName, " & _
		"com.MachineName, " & _
		"CASE " & _
			"WHEN Severity='-2147483648' THEN 'WARNING' " & _
			"WHEN Severity='-1073741824' THEN 'ERROR' " & _
			"ELSE 'INFO' " & _
			"END AS Category " & _
	"FROM " & _
		"dbo.v_StatusMessage stat " & _
			"JOIN v_ServerComponents com ON " & _
				"stat.SiteCode = com.SiteCode " & _
				"AND " & _
				"stat.MachineName = com.MachineName " & _
				"AND " & _
				"stat.Component = com.ComponentName " & _
	"WHERE " & _
		"Time > DATEADD(ss,-240-(24*3600),GetDate()) " & _
	"GROUP BY " & _
		"com.SiteCode, " & _
		"com.MachineName, " & _
		"com.ComponentName, " & _
		"Severity " & _
	") AS T1 " & _
	"GROUP BY " & _
		"T1.SiteCode, " & _
		"T1.ComponentName, " & _
		"T1.MachineName " & _
	"ORDER BY " & _
		"T1.ComponentName"
	
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_IntTableGrid rs, "", "", ""
CMWT_DB_CLOSE()

CMWT_SHOW_Query()
CMWT_Footer()
Response.Write "</body></html>"
%>