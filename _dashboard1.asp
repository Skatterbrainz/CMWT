<!-- begin-module: _dashboard1.asp -->
<%
'****************************************************************
' Filename..: _dashboard1.asp
' Author....: David M. Stein
' Date......: 11/27/2015
' Purpose...: home page
'****************************************************************

query1  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT ResourceID FROM dbo.v_R_System) AS T1"
query2  = "SELECT COUNT(*) AS QTY FROM dbo.v_R_System WHERE CLIENT0=1"
query3  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT User_Name0 FROM dbo.v_R_User WHERE Full_Domain_Name0='" & Application("CMWT_DOMAINSUFFIX") & "') AS T1"
query4  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT Name0 FROM dbo.v_R_UserGroup) AS T1"
query5  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT DisplayName0 FROM dbo.v_GS_ADD_REMOVE_PROGRAMS) AS T1"
query6  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT Full_Domain_Name0 FROM dbo.v_R_System) AS T1"
query7  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT DPname FROM dbo.vDPStatusPerDP) AS T1"
query8  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT ProductName FROM dbo.v_CM_AppDeployments) AS T1"
query9  = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT CollectionID FROM dbo.v_Collection) AS T1"
query10 = "SELECT COUNT(*) AS QTY FROM (SELECT DISTINCT AD_Site_Name0 FROM dbo.v_R_System) AS T1"
query11 = "SELECT COUNT(DISTINCT Name) AS QTY FROM dbo.vSMS_BoundaryGroup"
query12 = "SELECT COUNT(DISTINCT ResourceID) AS QTY FROM dbo.v_GS_SYSTEM_ENCLOSURE WHERE ChassisTypes0 IN (8,9,10,14,18)"
query13 = "SELECT COUNT(DISTINCT ResourceID) AS QTY FROM dbo.v_GS_SYSTEM_ENCLOSURE WHERE ChassisTypes0 IN (3,4,6,7,13,15,16)"

'----------------------------------------------------------------
Dim conn, cmd, rs

CMWT_DB_OPEN Application("DSN_CMDB")

count_computers = CMWT_DB_ROWCOUNT (query1)
count_clients   = CMWT_DB_ROWCOUNT (query2)
count_users     = CMWT_DB_ROWCOUNT (query3)
count_groups    = CMWT_DB_ROWCOUNT (query4)
count_apps      = CMWT_DB_ROWCOUNT (query5)
count_doms      = CMWT_DB_ROWCOUNT (query6)
count_dps       = CMWT_DB_ROWCOUNT (query7)
count_bgs       = CMWT_DB_ROWCOUNT (query11)
count_colls     = CMWT_DB_ROWCOUNT (query9)
count_sites     = CMWT_DB_ROWCOUNT (query10)
count_lt        = CMWT_DB_ROWCOUNT (query12)
count_dt        = CMWT_DB_ROWCOUNT (query13)

CMWT_DB_CLOSE()

'----------------------------------------------------------------

If count_clients > 0 And count_computers > 0 Then
	pct_clients = FormatPercent( count_clients / count_computers, 2)
	count_null = count_computers - count_clients
Else
	pct_clients = 0
	count_null = count_computers
End If
%>
<!-- end-module: _dashboard1.asp -->
