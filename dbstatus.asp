<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dbstatus.asp
' lastupdate..... 12/04/2016
' description.... database index fragmentation status
'-----------------------------------------------------------------------------
Response.Expires = -1
time1 = Timer
PageTitle    = "Database Index Fragmentation Status"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
blist = "<input type=""button"" name=""bdf"" id=""bdf"" value=""Defrag Now"" " & _
	"class=""btx w150 h32"" onClick=""document.location.href='dbdefrag.asp?x=yes'"" />"

Response.Write "<table class=""tfx"">"
query = "USE CM_" & Application("CM_SITECODE") & "; " & _
	"SELECT DISTINCT " & _
	"sch.name + '.' + OBJECT_NAME(stat.object_id) AS ITEMName, " & _
	"ind.name AS OBJ_NAME, CONVERT(int,stat.avg_fragmentation_in_percent) AS FRAG_PCT " & _
	"FROM sys.dm_db_index_physical_stats(DB_ID(),NULL,NULL,NULL,'LIMITED') stat " & _
	"JOIN sys.indexes ind ON stat.object_id=ind.object_id AND stat.index_id=ind.index_id " & _
	"JOIN sys.objects obj ON obj.object_id=stat.object_id " & _
	"JOIN sys.schemas sch ON obj.schema_id=sch.schema_id " & _
	"WHERE ind.name IS NOT NULL AND stat.avg_fragmentation_in_percent > 10.0 AND ind.type > 0 " & _
	"ORDER BY CONVERT(int,stat.avg_fragmentation_in_percent) DESC"
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query

If Not(rs.BOF And rs.EOF) Then
	xcols = rs.Fields.Count
	xrows = rs.RecordCount

	Response.Write "<tr>"
	For i = 0 to xcols-1
		Response.Write "<td class=""td6 v10 bgGray"">" & rs.Fields(i).Name & "</td>"
	Next
	Response.Write "</tr>"

	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value

			Select Case Ucase(fn)
				Case "FRAG_PCT":
					If fv >= 90 Then
						fv = "<span class=""cRed"">" & fv & "%</span>"
					ElseIf fv >= 80 Then
						fv = "<span class=""cOrange"">" & fv & "%</span>"
					ElseIf fv >= 40 Then
						fv = "<span class=""cYellow"">" & fv & "%</span>"
					Else
						fv = "<span class=""cLightGreen"">" & fv & "%</span>"
					End If
					Response.Write "<td class=""td6 v10 right"">" & fv & "</td>"
				Case Else:
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select

		Next
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	Response.Write "<tr>" & _
		"<td class=""td4 v10 bgGray"" colspan=""" & xcols & """>" & _
		xrows & " rows returned</td></tr>"

Else
	Response.Write "<tr class=""h100 tr1"">" & _
	"<td class=""td4 v10 ctr"">No matching record found</td></tr>"
End If

CMWT_DB_CLOSE()

Response.Write "</table>"

If CMWT_ADMIN() Then
	Response.Write "<div class=""tfx"">" & _
		"<br/><input type=""button"" name=""bDF"" id=""bDF"" value=""Defragment"" class=""btx w150 h30"" onClick=""document.location.href='dbdefrag.asp?x=yes'"" title=""Defragment Database Indexes"" /></div>"
End If

CMWT_Footer() 
Response.Write "</body></html>"
%>
