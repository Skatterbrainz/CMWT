<!-- #include file=_core.asp -->
<!-- #include file=_queries.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dpgroup.asp
' lastupdate..... 03/20/2016
' description.... distribution point group information
'-----------------------------------------------------------------------------
time1 = Timer

GroupNM = CMWT_GET("gn", "")
SortBy  = CMWT_GET("s", "ServerName")
QueryOn = CMWT_GET("qq", "")

CMWT_VALIDATE GroupNM, "DP Server Group Name was not specified"

pageTitle = "Distribution Point Group: " & GroupNM

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
	
query = "SELECT DISTINCT " & _
	"dbo.v_DistributionPoints.ServerName AS DPServer, " & _
	"dbo.v_DistributionPoints.Description AS DPComment, " & _
	"dbo.v_DistributionPoints.IsPeerDP AS Peer, " & _
	"dbo.v_DistributionPoints.IsPullDP AS PullDP, " & _
	"dbo.v_DistributionPoints.IsBITS AS BITS, " & _
	"dbo.v_DistributionPoints.IsMulticast AS MultiCast, " & _
	"dbo.v_DistributionPoints.IsProtected AS Protected, " & _
	"dbo.v_DistributionPoints.PreStagingAllowed AS Prestaged, " & _
	"dbo.v_DistributionPoints.IsPXE AS PXE, " & _
	"dbo.v_DistributionPoints.TransferRate, " & _
	"dbo.v_DistributionPoints.Priority " & _
	"FROM dbo.vSMS_DistributionPointGroup INNER JOIN " & _
	"dbo.v_DPGroupMembers ON dbo.vSMS_DistributionPointGroup.GroupID = dbo.v_DPGroupMembers.GroupID " & _
	"INNER JOIN " & _
	"dbo.v_DistributionPoints ON dbo.v_DPGroupMembers.DPNALPath = dbo.v_DistributionPoints.NALPath " & _
	"WHERE Name='" & GroupNM & "' " & _
	"ORDER BY " & SortBy
	
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "./?sbx1=1&sbx2=1&sbx4=dpgroups.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_Footer()
%>

</body>
</html>