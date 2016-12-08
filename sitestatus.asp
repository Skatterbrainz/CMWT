<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: sitestatus.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: site system status summary
'****************************************************************
time1 = Timer

PageTitle = "Site Status (Last 24 hours)"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
SortBy  = CMWT_GET("s", "Status")
QueryON = CMWT_GET("qq", "")

ffn = CMWT_GET("f", "")
ffv = CMWT_GET("v", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	Response.Write "<table class=""tfx"">"
	
		query = "SELECT * FROM " & _
				"(SELECT DISTINCT " & _
					"CASE " & _
						"WHEN Status=0 THEN 'GOOD' " & _
						"ELSE 'BAD' END AS StatusName, " & _
					"SUBSTRING(SiteSystem,13,(LEN(SiteSystem) - PATINDEX('%]%', SiteSystem))-26) AS ServerName, " & _
					"Role, " & _
					"BytesTotal, " & _
					"BytesFree, " & _
					"IIF(PercentFree < 0, 'Auto-Grow', CAST(PercentFree AS VARCHAR(10))) AS [PctFree], " & _
					"DownSince, " & _
					"TimeReported, " & _
					"CASE " & _
						"WHEN AvailabilityState=4 THEN 'ONLINE' " & _
						"WHEN AvailabilityState=0 THEN 'UNKNOWN' " & _
						"ELSE 'OTHER' END AS [State] " & _
					"FROM dbo.vSummarizer_SiteSystem WHERE TimeReported > '" & DateAdd("d", NOW, -1) & "') AS T1"
				
		If ffn <> "" and ffv <> "" Then
			query = query & " WHERE (T1." & ffn & "='" & ffv & "')"
			filtered = True
		End If
			
		Dim conn, cmd, rs
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		
		If Not(rs.BOF and rs.EOF) Then
			xrows = rs.RecordCount
			xcols = rs.Fields.Count
			
			Response.Write "<tr>"
			For i = 0 to xcols - 1
				fn = rs.Fields(i).Name
				Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
			Next
			Response.Write "</tr>"
			
			Do Until rs.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to xcols - 1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					If fn = "StatusName" Then
						Select Case fv
							Case "GOOD":
								tdc = "bgGreen"
							Case Else:
								tdc = "bgLightOrange"
						End Select
						fx = "<a href=""./?sbx1=2&sbx2=1&sbx4=sitestatus.asp&f=" & fn & "&v=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
						Response.Write "<td class=""td6 v10 " & tdc & " ctr"">" & fx & "</td>"
					Else
						fx = "<a href=""./?sbx1=2&sbx2=1&sbx4=sitestatus.asp&f=" & fn & "&v=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
						Response.Write "<td class=""td6 v10"">" & fx & "</td>"
					End If
				Next
				Response.Write "</tr>"
				rs.MoveNext
			Loop
			Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
				xrows & " rows returned"
			If filtered Then
				Response.Write " (filtered list - <a href=""./?sbx1=2&sbx2=1&sbx4=sitestatus.asp"" title=""Show All"">Show All</a>)"
			End If
			Response.Write "</td></tr>"
		Else
			Response.Write "<tr class=""h100""><td class=""td6 v10 ctr"">No rows returned</td></tr>"
		End If
		
		CMWT_DB_CLOSE()
		
	Response.Write "</table>"

	CMWT_SHOW_QUERY() 
	CMWT_Footer()
%>
	
</body>

</html>