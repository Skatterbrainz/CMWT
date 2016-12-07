<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... appvendors.asp
' lastupdate..... 11/30/2016
' description.... applications inventory by vendor
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle = "Applications by Vendor"
PageBackLink = "software.asp"
PageBackName = "Software"
SortBy  = CMWT_GET("s", "Products DESC")
QueryON = CMWT_GET("qq", "")
tcount  = CMWT_CM_APPCOUNT()

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
	
	<table class="tfx">
		<%
		Dim conn, cmd, rs
		
		query = "SELECT Vendor, Products FROM (" & _
			"SELECT DISTINCT Publisher0 AS Vendor, COUNT(DISTINCT DisplayName0) AS Products " & _
			"FROM dbo.v_GS_ADD_REMOVE_PROGRAMS " & _
			"WHERE Publisher0 IS NOT NULL AND LTRIM(Publisher0)<>'' " & _
			"GROUP BY Publisher0) AS T1 " & _
			"ORDER BY T1." & SortBy
				
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		
		If Not(rs.BOF and rs.EOF) Then
			xrows = rs.RecordCount
			xcols = rs.Fields.Count
			
			Response.Write "<tr>"
			For i = 0 to xcols - 1
				fn = rs.Fields(i).Name
				fx = CMWT_SORTLINK ("appvendors.asp", fn, SortBy)
				Response.Write "<td class=""td6 v10 bgGray"">" & fx & "</td>"
			Next
			Response.Write "</tr>"
			
			Do Until rs.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to xcols - 1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					Select Case Ucase(fn)
						Case "VENDOR":
							fv = "<a href=""vendorapps.asp?vn=" & fv & """ title=""Applications by " & fv & """>" & fv & "</a>"
					End Select
					Response.Write "<td class=""td6 v10"">" & fv & "</td>"
				Next
				Response.Write "</tr>"
				rs.MoveNext
			Loop
			Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
				xrows & " rows returned</td></tr>"
		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No rows returned</td></tr>"
		End If
		
		CMWT_DB_CLOSE()
		
		%>
	</table>
	
	<% 
	CMWT_SHOW_QUERY()
	CMWT_FOOTER() 
	%>
	
</body>

</html>