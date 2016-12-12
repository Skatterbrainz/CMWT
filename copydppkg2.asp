<!-- #include file="_core.asp" -->
<%
'-----------------------------------------------------------------------------
' filename....... copydppkg2.asp
' lastupdate..... 12/12/2016
' description.... 
'-----------------------------------------------------------------------------
time1 = Timer

sourceDP = CMWT_GET("s1", "")
targetDP = CMWT_GET("s2", "")

If sourceDP = "" And targetDP = "" Then
	CMWT_STOP "Source and Target DP server names must be specified"
End If

If sourceDP = targetDP Then
	CMWT_STOP "Source and Target DP cannot be the same."
End If

Const wbemFlagReturnImmediately = &H10 
Const wbemFlagForwardOnly = &H20 

self_link = "dppackages.asp"

query = "SELECT DISTINCT TOP 10 " & _
		"dbo.v_Package.Name AS PackageName, " & _
		"dbo.v_DistributionPoint.LastRefreshTime AS DPRefreshDate, " & _
		"dbo.v_Package.SourceSite AS SiteCode, " & _
		"dbo.v_Package.PackageID, " & _
		"dbo.v_PackageStatusRootSummarizer.Targeted, " & _
		"dbo.v_PackageStatusRootSummarizer.Installed, " & _
		"dbo.v_PackageStatusRootSummarizer.Failed, " & _
		"dbo.v_PackageStatusRootSummarizer.SourceSize " & _
	"FROM dbo.v_Package " & _
		"INNER JOIN " & _
		"dbo.v_DistributionPoint ON dbo.v_Package.PackageID = dbo.v_DistributionPoint.PackageID " & _
		"INNER JOIN " & _
		"dbo.v_PackageStatusRootSummarizer ON dbo.v_Package.PackageID = dbo.v_PackageStatusRootSummarizer.PackageID " & _
	"WHERE " & _
		"(dbo.v_DistributionPoint.ServerNALPath LIKE '%" & sourceDP & "%') " & _
		"AND " & _
		"(dbo.v_DistributionPoint.ServerNALPath NOT LIKE '%" & targetDP & "%') " & _
	"ORDER BY PackageName"

Set conn = Server.CreateObject("ADODB.Connection")
Set cmd  = Server.CreateObject("ADODB.Command")
Set rs   = Server.CreateObject("ADODB.Recordset")

conn.Open Application("DSN_CMDB")

rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockReadOnly

Set cmd.ActiveConnection = conn

cmd.CommandType = adCmdText
cmd.CommandText = query
rs.Open cmd

If Not(rs.BOF And rs.EOF) Then
	xcols = rs.Fields.Count
	xrows = rs.RecordCount
	found = True

	Response.Write "<br/>" & xrows & " packages were found on " & sourceDP & " which are not assigned to " & targetDP
	
	On Error Resume Next
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")
	ErrTrap "swbemLocator"

	Set objSSconn = objLocator.ConnectServer(".", "Root\SMS\Site_" & Application("CMWT_SITECODE"))

	Do Until rs.EOF
		pkgID = rs.Fields("PackageID").value
		pkgNM = rs.Fields("PackageName").value
		
		Response.Write "<br/>copied package: " & pkgID & " - " & pkgNM
		rs.MoveNext
	Loop
	
	rs.Close
	conn.Close
	
	objSSconn.Close
	Set objSSConn = Nothing
	Set objLocator = Nothing
	
Else
	Response.Write "<p>No packages were found on " & sourceDP
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	Response.End
End If

Response.Flush
Response.Redirect "dppackages.asp?id=" & sourceDP

'----------------------------------------------------------------
' function: 
'----------------------------------------------------------------

Function CMWT_DP_AddPackage (objSS, pkgID, targetServer)
	Dim objDP, query, colResources
	Dim objResource, rescount, result
	
	result = 0
	
	On Error Resume Next
	
	Set objDP = objSS.Get("SMS_DistributionPoint").SpawnInstance_
	
	If err.Number <> 0 Then
		result = err.Number
		CMWT_DP_AddPackage = result
		Exit Function
	End If
	
	objDP.PackageID = pkgID     
	
	query = "SELECT * FROM SMS_SystemResourceList " & _
		"WHERE RoleName='SMS Distribution Point' " & _
		"AND SiteCode='" & Application("CMWT_SITECODE") & "' " & _
		"AND ServerName='" & targetServer & "' " & _
		"AND NALPath NOT LIKE '%PXE%'"
	
	Set colResources = objSS.ExecQuery(query, , wbemFlagForwardOnly Or wbemFlagReturnImmediately)
	
	If err.Number <> 0 Then
		result = err.Number
		Set objDP = Nothing
		CMWT_DP_AddPackage = result
		Exit Function
	End If
	
	rescount = 0
	
	For Each objResource In colResources      
		objDP.ServerNALPath = objResource.NALPath
		objDP.SiteCode = objResource.SiteCode        
		rescount = rescount + 1
	Next
	
	If err.Number <> 0 Then
		result = err.Number
		Set objDP = Nothing
		Set colResources = Nothing
		CMWT_DP_AddPackage = result
		Exit Function
	Else
		If rescount = 1 Then
			objDP.Put_
			If err.Number = 0 Then
				result = 0
			End If
		Else
			result = -1
		End If
	End If
	
	CMWT_DP_AddPackage = result
	
End Function

'----------------------------------------------------------------
%>