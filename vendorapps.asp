<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... vendorapps.asp
' lastupdate..... 11/27/2016
' description.... applications installs for a specific vendor
'-----------------------------------------------------------------------------
time1 = Timer

VendorName = CMWT_GET("vn", "")
CMWT_VALIDATE VendorName, "Vendor Name was not specified"
pageTitle = "Applications: " & VendorName
SortBy  = CMWT_GET("s", "QTY DESC")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
	
	<table class="tfx">
		<%
		Dim conn, cmd, rs
		
		query = "SELECT ProductName, Vendor, QTY FROM (" & _
			"SELECT DISTINCT DisplayName0 AS ProductName, Publisher0 AS Vendor, COUNT(*) AS QTY " & _
			"FROM dbo.v_GS_ADD_REMOVE_PROGRAMS " & _
			"WHERE Publisher0 ='" & VendorName & "' " & _
			"GROUP BY DisplayName0, Publisher0, ResourceID " & _
			") AS T1 " & _
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
						Case "PRODUCTNAME":
							fv = "<a href=""app.asp?pn=" & fv & """ title=""Installations of " & fv & """>" & fv & "</a>"
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
	CMWT_Footer()
	%>
	
</body>

</html>
