<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... devices.asp
' lastupdate..... 12/07/2016
' description.... devices listing page
'-----------------------------------------------------------------------------
time1 = Timer
KeySet  = CMWT_GET("ks", "1")
QueryOn = CMWT_GET("qq", "")
ObjPfx  = CMWT_GET("ch", "C")
SortBy  = CMWT_GET("s", "DeviceName")

Select Case KeySet
	Case "1": PageTitle = "Devices"
	Case "2": PageTitle = "Devices: Servers"
	Case "3": PageTitle = "Devices: Clients"
	Case "4": PageTitle = "Devices: Desktops"
	Case "5": PageTitle = "Devices: Laptops"
End Select
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
CMWT_CLICKBAR objPfx, "devices.asp?ks=" & KeySet & "&ch="

Select Case KeySet
	Case "2"
		query = "SELECT Name0 AS DeviceName, ResourceID, " &_
			"AD_Site_Name0 AS ADSiteName," & _
			"Client_Version0 AS ClientVersion, " & _
			"CPUType0 AS CPUType, " & _
			"Creation_Date0 AS DateCreated, " & _
			"Operating_System_Name_and0 AS OSName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System " & _
			"Where (ResourceType = 5) " & _
			"AND (Operating_System_Name_and0 LIKE '%SERVER%')"
	Case "3"
		query = "SELECT Name0 AS DeviceName, ResourceID, " &_
			"AD_Site_Name0 AS ADSiteName," & _
			"Client_Version0 AS ClientVersion, " & _
			"CPUType0 AS CPUType, " & _
			"Creation_Date0 AS DateCreated, " & _
			"Operating_System_Name_and0 AS OSName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System " & _
			"Where (ResourceType = 5) " & _
			"AND (Operating_System_Name_and0 NOT LIKE '%SERVER%')"
	Case "4"
		query = "SELECT DISTINCT Name0 as DeviceName, " & _
			"AD_Site_Name0 AS ADSiteName, " & _
			"Client_Version0 AS ClientVersion, " & _
			"CPUType0 AS CPUType, " & _
			"Creation_Date0 AS DateCreated, " & _
			"Operating_System_Name_and0 AS OSName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM (" & q_devices & ") AS T1 " & _
			"WHERE (T1.ChassisTypes0 IN (3,4,6,7))"
	Case "5"
		query = "SELECT DISTINCT Name0 as DeviceName, " & _
			"AD_Site_Name0 AS ADSiteName, " & _
			"Client_Version0 AS ClientVersion, " & _
			"CPUType0 AS CPUType, " & _
			"Creation_Date0 AS DateCreated, " & _
			"Operating_System_Name_and0 AS OSName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM (" & q_devices & ") AS T1 " & _
			"WHERE (T1.ChassisTypes0 IN (9,10,14))"
	Case Else
		query = "SELECT Name0 AS DeviceName, ResourceID, " &_
			"AD_Site_Name0 AS ADSiteName," & _
			"Client_Version0 AS ClientVersion, " & _
			"CPUType0 AS CPUType, " & _
			"Creation_Date0 AS DateCreated, " & _
			"Operating_System_Name_and0 AS OSName, " & _
			"Virtual_Machine_Host_Name0 AS VMHost " & _
			"FROM dbo.v_R_System " & _
			"Where (ResourceType = 5)"
End Select

If objPFX <> "ALL" Then
	query = query & " AND (Name0 LIKE '" & ObjPfx & "%')"
End If

query = query & " ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query

Response.Write "<form name=""form1"" id=""form1"" method=""post"" action=""cmcx.asp"">" & _
	"<input type=""hidden"" name=""mx"" id=""mx"" value=""ADD"" />"

If Not (rs.BOF and rs.EOF) Then 
	xrows = rs.RecordCount 
	xcols = rs.Fields.Count
	Response.Write "<table class=""tfx""><tr>"
	' display column headings
	Response.Write "<td class=""td6 v10 ctr w30 bgGray"">&nbsp;</td>"
	For i = 0 to xcols -1
		fn = rs.Fields(i).Name
		Select Case Ucase(fn)
			Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
				Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
			Case Else:
				Response.Write "<td class=""td6 v10 bgGray"">"
		End Select
		Response.Write CMWT_SORTLINK("devices.asp?ch=" & objPFX, fn, SortBy) & "</td>"
	Next
	Response.Write "</tr>"
	' iterate dataset rows
	afn = ""
	flx = Split("cn=devicename", "=")
	' form property name
	fpn = flx(0)
	' form recordset column name
	fcn = flx(1)

	Do Until rs.EOF
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10 ctr"">"
		If rs.Fields("ClientVersion").Value <> "" Then
			Response.Write "<input type=""checkbox"" class=""CB1"" name=""" & fpn & """ id=""" & _
			fpn & """ value=""" & rs.Fields(fcn).value & """ />"
		End If
		Response.Write "</td>"
		For i = 0 to xcols-1
			fn = rs.Fields(i).Name
			fv = rs.Fields(i).Value
			If Ucase(afn) = Ucase(fn) Then
				fv = "<a href=""" & afl & "=" & fv & """>" & fv & "</a>"
			Else
				fv = CMWT_AutoLink (fn, fv)
			End If
			Response.Write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
		next
		rs.MoveNext
	Loop
	Response.Write "<tr>" & _
		"<td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & xrows & " rows returned</td></tr></table>"
End If

'CMWT_DB_TableGrid2 rs, "", "devices.asp", "", "cn=devicename"

Response.Write "<table class=""tfx""><tr><td class=""v10 pad6"">" & _
	"<input type=""button"" name=""b0"" id=""b0"" class=""btx w140 h30"" value=""Clear All"" " & _
	"onClick=""document.location.href='devices.asp?ks=" & KeySet & "&ch=" & objPfx & "'"" title=""Clear All"" />&nbsp;" & _
	"<select name=""cid"" id=""cid"" size=""1"" class=""pad5 h30 w300"">" & _
	"<option value=""""></option>"
	CMWT_CM_ListCollections conn, "", 2, ""
Response.Write "</select>" & _
	"&nbsp;<input type=""submit"" name=""b1"" id=""b1"" class=""btx w140 h30"" " & _
	"value=""Add Members"" title=""Add Selected Members to Collection"" />" & _
	" (only Direct-Membership Collections can be modified)</td></tr></table></form>"

CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>