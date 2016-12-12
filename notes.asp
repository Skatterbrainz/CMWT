<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... notes.asp
' lastupdate..... 12/03/2016
' description.... notes library report
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle    = "Notes Library"
PageBackLink = "admin.asp"
PageBackName = "Administration"
SelfLink  = "notes.asp"
SortBy    = CMWT_GET("s", "AttachClass")
QueryON   = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
	
<%
Response.Write "<table class=""tfx"">"
query = "SELECT NoteID, AttachClass, AttachedTo, Comment, DateCreated, CreatedBy " & _
	"FROM dbo.Notes " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMWT"), query

If Not(rs.BOF And rs.EOF) Then
	found = True
	xrows = rs.RecordCount
	xcols = rs.Fields.Count

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
				Case "NOTEID":
					fv = CMWT_IMG_LINK (TRUE, "icon_del2", "icon_del1", "icon_del3", "confirm.asp?id=" & fv & "&tn=notes&pk=noteid&t=notes.asp", "Remove") & " " & _
						CMWT_IMG_LINK (TRUE, "icon_edit2", "icon_edit1", "icon_edit2", "noteedit.asp?id=" & fv, "Edit")
					Response.Write "<td class=""td6 v10 w50"">" & fv & "</td>"
				Case Else:
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
			End Select

		Next
		Response.Write "</tr>"
		rs.MoveNext
	Loop
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
		xrows & " rows returned</td></tr>"
Else
	Response.Write "<tr class=""h100 tr1"">" & _
		"<td class=""td6 v10 ctr"">No matching rows returned</td></tr>"
End If

CMWT_DB_CLOSE()

Response.Write "</table>"
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>
	
</body>
</html>
