<!-- #include file=_core.asp -->
<%
'****************************************************************
' Filename..: clientpush.asp
' Author....: David M. Stein
' Date......: 11/30/2016
' Purpose...: client push installations
'****************************************************************
time1 = Timer

PageTitle = "Client Push Installations"
PageBackLink = "reports.asp"
PageBackName = "Reports"
SortBy  = CMWT_GET("s", "DeviceName")
QueryON = CMWT_GET("qq", "")

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT DISTINCT " & _
	"T2.Name AS DeviceName, " & _
	"T1.AD_Site_Name0 AS ADSiteName, " & _
	"T1.Client_Version0 AS ClientVer, " & _
	"T2.Forced, " & _
	"T2.ForceReinstall AS Reinstall, " & _
	"T2.PushEvenIfDC AS PushToDC, " & _
	"T2.AssignedSiteCode AS Assigned, " & _
	"T2.InitialRequestDate AS Requested, " & _
	"T2.LatestProcessingAttempt AS Latest, " & _
	"T2.LastErrorCode, " & _
	"T2.NumProcessAttempts AS Attempts " & _
	"FROM dbo.v_R_System T1 RIGHT JOIN " & _
	"dbo.ClientPushMachine_G T2 ON " & _
	"T1.ResourceID = T2.MachineID " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "clientpush.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>