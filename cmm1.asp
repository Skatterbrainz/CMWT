<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmm1.asp
' lastupdate..... 12/29/2016
' description.... cmmonitor database commands table summary
'-----------------------------------------------------------------------------
time1 = Timer

if Application("DSN_CMM") = "" Then
	CMWT_STOP "CM Monitor connection is not configured in _config.txt"
end if

SortBy    = CMWT_GET("s", "ID DESC")
KeySet    = CMWT_GET("x", "1")
FilterFN  = CMWT_GET("fn", "")
FilterFV  = CMWT_GET("fv", "")
QueryOn   = CMWT_GET("qq", "")

PageTitle    = "CMMonitor Commands"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
if FilterFN <> "" Then
	filtered = True
end if

Response.Write "<table class=""t2""><tr>"

If filtered <> True Then
	Select Case KeySet
		Case "1"
			Response.Write "<td class=""m22"">Top 50</td>"
			Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=2&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 100</td>"
		Case "2"
			Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=1&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 50</td>"
			Response.Write "<td class=""m22"">Top 100</td>"
		Case Else
			Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=1&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 50</td>"
			Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=2&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 100</td>"
			Response.Write "<td class=""m22"">Filtered</td>"
	End Select
Else
	Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=1&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 50</td>"
	Response.Write "<td class=""m11"" onClick=""document.location.href='cmm1.asp?x=2&fn=" & FilterFN & "&fv=" & FilterFV & "'"">Top 100</td>"
	Response.Write "<td class=""m22"">Filtered</td>"
End If

Response.Write "</tr></table>"

Dim conn, cmd, rs

Select Case Ucase(FilterFN)
	Case "ID"

		query = "SELECT TOP 1 " & _
			"ID,DatabaseName,SchemaName,ObjectName,ObjectType, " & _
			"IndexName,IndexType,StatisticsName,PartitionNumber, " & _
			"ExtendedInfo,Command,CommandType,StartTime,EndTime, " & _
			"ErrorNumber,ErrorMessage " & _
			"FROM dbo.CommandLog " & _
			"WHERE (ID=" & FilterFV & ")"
		CMWT_DB_QUERY Application("DSN_CMM"), query
		CMWT_DB_TABLEROWGRIDFilter rs, "", "cmm1.asp"
	
	Case ""

		if KeySet = "1" Then
			query = "SELECT TOP 50 " & _
				"ID,DatabaseName,SchemaName,ObjectName,ObjectType, " & _
				"CommandType,StartTime,EndTime " & _
				"FROM dbo.CommandLog " & _
				"ORDER BY " & SortBy
		else
			query = "SELECT TOP 100 " & _
				"ID,DatabaseName,SchemaName,ObjectName,ObjectType, " & _
				"CommandType,StartTime,EndTime " & _
				"FROM dbo.CommandLog " & _
				"ORDER BY " & SortBy
		end if
		'CMWT_DB_TableGridFilter [rs, Caption, SortLink, AutoLink, ColumnSet, FilterLink
		CMWT_DB_QUERY Application("DSN_CMM"), query
		CMWT_DB_TableGridFilter rs, "", "cmm1.asp", "", "", "cmm1.asp"
	
	Case Else

		if KeySet = "1" Then 
			query = "SELECT TOP 50 " & _
				"ID,DatabaseName,SchemaName,ObjectName,ObjectType, " & _
				"CommandType,StartTime,EndTime " & _
				"FROM dbo.CommandLog " & _
				"WHERE (" & FilterFN & "='" & FilterFV & "') " & _
				"ORDER BY " & SortBy
		else
			query = "SELECT TOP 100 " & _
				"ID,DatabaseName,SchemaName,ObjectName,ObjectType, " & _
				"CommandType,StartTime,EndTime " & _
				"FROM dbo.CommandLog " & _
				"WHERE (" & FilterFN & "='" & FilterFV & "') " & _
				"ORDER BY " & SortBy
		end if
		CMWT_DB_QUERY Application("DSN_CMM"), query
		CMWT_DB_TableGridFilter rs, "", "cmm1.asp", "", "", "cmm1.asp"

End Select

CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

Response.Write "</body></html>"
%>
