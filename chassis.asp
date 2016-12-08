<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... chassis.asp
' lastupdate..... 11/30/2016
' description.... computers by chassis type (form factor)
'-----------------------------------------------------------------------------
time1 = Timer

SortBy  = CMWT_GET("s", "ChassisType")
ChassTp = CMWT_GET("t", "")
QueryOn = CMWT_GET("qq", "")
PageTitle = "Devices by Chassis Type"
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx"">"

query = "SELECT ChassisType, QTY " & _
	"FROM (" & _
	"SELECT DISTINCT ChassisTypes0, " & _
	"CASE " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '1' THEN 'Virtual' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '2' THEN 'Blade Server' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '3' THEN 'Desktop' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '4' THEN 'Low-Profile Desktop' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '5' THEN 'Pizza-Box' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '6' THEN 'Mini Tower' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '7' THEN 'Tower' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '8' THEN 'Portable' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '9' THEN 'Laptop' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '10' THEN 'Notebook' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '11' THEN 'Hand-Held'" & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '12' THEN 'Mobile Device in Docking Station' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '13' THEN 'All-in-One' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '14' THEN 'Sub-Notebook' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '15' THEN 'Space Saving Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '16' THEN 'Ultra Small Form Factor' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '17' THEN 'Server Tower Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '18' THEN 'Mobile Device in Docking Station' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '19' THEN 'Sub-Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '20' THEN 'Bus-Expansion chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '21' THEN 'Peripheral Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '22' THEN 'Storage Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '23' THEN 'Rack-Mounted Chassis' " & _
	"WHEN dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0 LIKE '24' THEN 'Sealed-Case PC' " & _
	"ELSE 'Unknown' " & _
	"END AS 'ChassisType'," & _
	"COUNT(DISTINCT ResourceID) AS QTY " & _
	"FROM dbo.v_GS_SYSTEM_ENCLOSURE " & _
	"XXX " & _
	"GROUP BY ChassisTypes0 ) AS T1"

If ChassTP <> "" Then
	query = Replace(query, "XXX", "WHERE CHASSISTYPES0 IN (" & ChassTp & ")")
	filtered = True
Else
	query = Replace(query, "XXX ", "")
End If
query = query & " ORDER BY " & SortBy
		
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "chassis.asp", ""
CMWT_DB_CLOSE()

If filtered Then
	Response.Write "<tr>" & _
		"<td class=""td6 bgGray"" colspan=""2"">" & _
		"Filtered Results : <a href=""chassis.asp"" title=""Remove Filter"">Remove Filter</a></td></tr>"
End If 
Response.Write "</table>"

CMWT_SHOW_QUERY() 
CMWT_Footer()
%>
	
</body>
</html>