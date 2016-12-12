<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... search.asp
' lastupdate..... 12/11/2016
' description.... search results report
'-----------------------------------------------------------------------------
time1 = Timer

SearchVal = CMWT_GET("q", "")
SearchCat = CMWT_GET("cat", "")
CMWT_VALIDATE SearchVal, "No search value was entered"
QueryOn  = CMWT_GET("qq", "")

pageTitle = "Search Results: " & SearchVal

qlist = "v_R_System=Name0=AD Computers=1|" & _
	"v_R_User=Name0=AD Users=2|" & _
	"v_R_UserGroup=Usergroup_Name0=AD Groups=3|" & _
	"v_Collection=Name=Collections=4|" & _
	"v_Package=Name=Packages=5|" & _
	"v_GS_COMPUTER_SYSTEM=Model0=Computer Models=6|" & _
	"v_GS_SoftwareProduct=ProductName=Installed Software=7|" & _
	"v_DriverPackage=Name=Driver Packages=8"

' number=category=column=fieldslist=tablename|
	
xlist = "1=AD Computers=Name0=Name0 AS ComputerName,ResourceID,AD_Site_Name0 AS ADSite,Client_Version0 AS ClientVersion,Virtual_Machine_Host_Name0 AS VMHost=v_R_System|" & _
	"2=AD Users=Name0=Full_User_Name0 AS DisplayName,User_Name0 AS UserID,User_Principal_Name0 AS UPN,Windows_NT_Domain0 AS Domain=v_R_User|" & _
	"3=AD Groups=Usergroup_Name0=Usergroup_Name0 AS GroupName,Windows_NT_Domain0 AS Domain=v_R_UserGroup|" & _
	"4=Collections=Name=Name,CollectionID,Comment,CASE WHEN CollectionType > 1 THEN 'USERS' ELSE 'DEVICES' END AS CollType,MemberCount=v_Collection|" & _
	"5=Packages=Name=Name,PackageID,Manufacturer,Description=v_Package|" & _
	"6=Computer Models=Model0=Manufacturer0 AS Mfr,Model0 AS Model,ResourceID,Domain0 AS Domain,Name0 AS Name,SystemType0 AS SystemType=v_GS_COMPUTER_SYSTEM|" & _
	"7=Installed Software=ARPDisplayName0=ARPDisplayName0 AS ProductName,NormalizedPublisher AS Publisher,NormalizedVersion AS ProductVersion=v_GS_INSTALLED_SOFTWARE_CATEGORIZED|" & _
	"8=Driver Packages=Name=Name,PackageID,Description=v_DriverPackage"
	
Function CMWT_SEARCH_COUNT (c, SearchVal, TableName, MatchField)
	Dim query, cmd, rs, result : result = 0
	query = "SELECT COUNT(*) AS QTY FROM dbo." & TableName & _
		" WHERE " & MatchField & " LIKE '%" & SearchVal & "%'"
	Set cmd  = CreateObject("ADODB.Command")
	Set rs   = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = rs.Fields("QTY").value
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_SEARCH_COUNT = result
End Function

Function CMWT_GET_SEARCH_CAT_QUERY (vList, vCat)
	Dim x, y, result : result = ""
	For each x in Split(vList, "|")
		y = Split(x,"=")
		if y(0) = vCat Then
			result = "SELECT DISTINCT " & y(3) & _
				" FROM dbo." & y(4) & _
				" WHERE (" & y(2) & " LIKE '%XXXX%')"
		End If
	Next
	CMWT_GET_SEARCH_CAT_QUERY = result
End Function 

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Dim conn, cmd, rs

If SearchCat = "" Then
	Response.Write "<table class=""tfx"">" & _
		"<tr><td class=""td6 v10 w50 ctr bgGray"">Hits</td>" & _
		"<td class=""td6 v10 bgGray"">Search Category</td></tr>"

	CMWT_DB_OPEN Application("DSN_CMDB")
			
	For each qset in Split(qlist, "|")
		qm  = Split(qset, "=")
		qty = CMWT_SEARCH_COUNT(conn, SearchVal, qm(0), qm(1))
		cat = qm(3)
		if qty > 0 Then 
			fv = "<a href=""search.asp?q=" & SearchVal & "&cat=" & cat & """>" & qm(2) & "</a>"
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10 w50 ctr bgGreen"">" & qty & "</td>" & _
				"<td class=""td6 v10"">" & fv & "</td></tr>"
		Else 
			fv = qm(2)
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10 w50 ctr"">" & qty & "</td>" & _
				"<td class=""td6 v10"">" & fv & "</td></tr>"
		End If
	Next
	
	Response.Write "</table>"
	
	CMWT_DB_CLOSE()
	query = "(query statement is embedded in function)"
	CMWT_SHOW_QUERY()
	
Else 
	
	query = Replace(CMWT_GET_SEARCH_CAT_QUERY(xlist, SearchCat), "XXXX", SearchVal)
	
	CMWT_DB_QUERY Application("DSN_CMDB"), query
	CMWT_DB_TABLEGRID rs, "", "", ""
	CMWT_DB_CLOSE()

	CMWT_SHOW_QUERY()
	
End If 

CMWT_Footer()
Response.Write "</body></html>"
%>
