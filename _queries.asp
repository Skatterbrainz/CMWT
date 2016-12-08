<%
'****************************************************************
' Filename..: _sccm_queries.asp
' Author....: David M. Stein
' Date......: 06/22/2014
' Purpose...: raw sql queries for sccm site database reporting
'****************************************************************

q_client_subnets = "SELECT DISTINCT " & _
	"dbo.v_R_System.Name0 AS ComputerName, dbo.v_R_System.ResourceID, dbo.v_GS_COMPUTER_SYSTEM.Model0 AS Model, " & _
	"dbo.v_R_System.AD_Site_Name0 AS ADSiteName, dbo.v_GS_OPERATING_SYSTEM.Caption0 AS Windows, " & _
	"dbo.v_GS_OPERATING_SYSTEM.CSDVersion0 AS ServicePack, dbo.v_R_System.Client0 AS Client, " & _
	"dbo.v_R_System.Full_Domain_Name0 AS DomainName " & _
	"FROM dbo.v_R_System INNER JOIN " & _
	"dbo.v_RA_System_IPSubnets ON dbo.v_R_System.ResourceID = dbo.v_RA_System_IPSubnets.ResourceID LEFT OUTER JOIN " & _
	"dbo.v_GS_COMPUTER_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_COMPUTER_SYSTEM.ResourceID LEFT OUTER JOIN " & _
	"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID"

q_site_config = "SELECT SiteCode,RoleID,RoleName,State,Configuration," & _
	"MessageID,LastEvaluatingTime,Param1,Param2,Param3,Param4,Param5,Param6 " & _
	"FROM dbo.vCM_SiteConfiguration"

q_boundaries = "SELECT BoundaryID,DisplayName,BoundaryType, " & _
	"CASE WHEN BoundaryType=1 THEN 'Site' WHEN BoundaryType=3 THEN 'IP Range' ELSE '' END AS BoundaryTypeName, " & _
	"BoundaryFlags, Value, GroupCount " & _
	"FROM dbo.vSMS_Boundary"

q_boundary_groups = "SELECT Name AS BoundaryGroup,GroupID,Description,DefaultSiteCode,MemberCount," & _
	"SiteSystemCount,Shared " & _
	"FROM dbo.vSMS_BoundaryGroup"

q_boundary_dpservers = "SELECT dbo.vSMS_BoundaryGroup.GroupID, dbo.vSMS_BoundaryGroup.Name AS GroupName, " & _
	"dbo.vSMS_BoundaryGroup.Description, dbo.vSMS_BoundaryGroup.DefaultSiteCode AS SiteCode, " & _
	"dbo.vSMS_BoundaryGroup.MemberCount, dbo.vDistributionPoints.ServerName AS DPServerName, " & _
	"dbo.vDistributionPoints.Description AS DPComment, dbo.vDistributionPoints.IsPeerDP, " & _
	"dbo.vDistributionPoints.IsPullDP, dbo.vDistributionPoints.IsPullDPInstalled, dbo.vDistributionPoints.IsFileStreaming, " & _
	"dbo.vDistributionPoints.IsBITS, dbo.vDistributionPoints.IsMulticast, dbo.vDistributionPoints.AnonymousEnabled, " & _
	"dbo.vDistributionPoints.DPType, dbo.vDistributionPoints.IsProtected, dbo.vDistributionPoints.PreStagingAllowed, " & _
	"dbo.vDistributionPoints.IsPXE, dbo.vDistributionPoints.State " & _
	"FROM dbo.vSMS_BoundaryGroup INNER JOIN " & _
	"dbo.vSMS_BoundaryGroupSiteSystems ON dbo.vSMS_BoundaryGroup.GroupID = dbo.vSMS_BoundaryGroupSiteSystems.GroupID " & _
	"INNER JOIN dbo.vDistributionPoints ON dbo.vSMS_BoundaryGroupSiteSystems.ServerNALPath = dbo.vDistributionPoints.NALPath"

q_bg_members = "SELECT DISTINCT dbo.vSMS_Boundary.DisplayName, dbo.vSMS_Boundary.BoundaryID, dbo.vSMS_BoundaryGroup.Name, " & _
	"dbo.vSMS_BoundaryGroup.GroupID, dbo.vSMS_Boundary.Value " & _
	"FROM dbo.vSMS_BoundaryGroupMembers INNER JOIN " & _
	"dbo.vSMS_BoundaryGroup ON dbo.vSMS_BoundaryGroupMembers.GroupID = dbo.vSMS_BoundaryGroup.GroupID INNER JOIN " & _
	"dbo.vSMS_Boundary ON dbo.vSMS_BoundaryGroupMembers.BoundaryID = dbo.vSMS_Boundary.BoundaryID"

q_dist_status = "SELECT SoftwareName,Installed,Retrying,Failed," & _
	"LastUpdated,PkgID,AppCI " & _
	"FROM dbo.vDPStatus"

q_deployment_status = "SELECT CollectionID,AssignmentID,CollectionName,SoftwareName,FeatureType," & _
	"CASE WHEN FeatureType=1 THEN 'Application' WHEN FeatureType=5 THEN 'Package' ELSE '' END AS PayloadType," & _
	"SummaryType,DeploymentIntent," & _
	"CASE WHEN DeploymentIntent=1 THEN 'Required' WHEN DeploymentIntent=2 THEN 'Available' ELSE '' END AS DeploymentTypeName," & _
	"EnforcementDeadline,NumberSuccess,NumberInProgress,NumberUnknown,NumberErrors,NumberOther," & _
	"NumberTotal,SummarizationTime,ProgramName,PackageID " & _
	"FROM dbo.vDeploymentSummary"

q_dp_status = "SELECT DISTINCT REPLACE(DPName, '." & DomSuffix & "', '') AS ServerName, DPName AS FullName, " & _
	"Installed, Retrying, Failed, Installed + Failed + Retrying AS iTotal, LastUpdated " & _
	"FROM dbo.vDPStatusPerDP WHERE (LTRIM(SoftwareName) <> '')"

q_disc_forests = "SELECT Description,ForestFQDN,DiscoveryEnabled,ForestID,PublishingEnabled," & _
	"Tombstoned,LastDiscoveryTime,LastDiscoveryStatus,PublishingStatus,DiscoveredTrusts," & _
	"DiscoveredDomains,DiscoveredADSites,DiscoveredIPSubnets " & _
	"FROM dbo.vActiveDirectoryForests"

q_site_systems = "SELECT DISTINCT dbo.v_BoundarySiteSystems.SiteSystemName, dbo.v_SiteSystemSummarizer.SiteSystem, " & _
	"dbo.v_SiteSystemSummarizer.SiteCode, dbo.v_SiteSystemSummarizer.Role, dbo.v_SiteSystemSummarizer.Status, " & _
	"dbo.v_SiteSystemSummarizer.BytesTotal, dbo.v_SiteSystemSummarizer.BytesFree, dbo.v_SiteSystemSummarizer.PercentFree, " & _
	"dbo.v_SiteSystemSummarizer.DownSince, dbo.v_SiteSystemSummarizer.TimeReported, dbo.v_SiteSystemSummarizer.AvailabilityState " & _
	"FROM dbo.v_SiteSystemSummarizer INNER JOIN " & _
	"dbo.v_BoundarySiteSystems ON dbo.v_SiteSystemSummarizer.SiteSystem = dbo.v_BoundarySiteSystems.ServerNALPath"

q_printers = "SELECT DISTINCT dbo.v_GS_PRINTER_DEVICE.ResourceID, dbo.v_GS_PRINTER_DEVICE.Name0, dbo.v_GS_PRINTER_DEVICE.Description0, " & _
	"dbo.v_GS_PRINTER_DEVICE.DeviceID0, dbo.v_GS_PRINTER_DEVICE.DriverName0, dbo.v_GS_PRINTER_DEVICE.ShareName0, " & _
	"dbo.v_GS_PRINTER_DEVICE.Status0, dbo.v_R_System.Name0 AS ComputerName, dbo.v_R_System.AD_Site_Name0, dbo.v_R_System.Full_Domain_Name0 " & _
	"FROM dbo.v_GS_PRINTER_DEVICE INNER JOIN " & _
	"dbo.v_R_System ON dbo.v_GS_PRINTER_DEVICE.ResourceID = dbo.v_R_System.ResourceID"

q_installed_apps = "SELECT DISTINCT dbo.v_R_System.Name0, dbo.v_R_System.ResourceID, dbo.v_GS_ADD_REMOVE_PROGRAMS.DisplayName0, " & _
	"dbo.v_GS_ADD_REMOVE_PROGRAMS.Publisher0, dbo.v_GS_ADD_REMOVE_PROGRAMS.InstallDate0, dbo.v_R_System.AD_Site_Name0, " & _
	"dbo.v_R_System.Operating_System_Name_and0 " & _
	"FROM dbo.v_R_System LEFT OUTER JOIN " & _
	"dbo.v_GS_ADD_REMOVE_PROGRAMS ON dbo.v_R_System.ResourceID = dbo.v_GS_ADD_REMOVE_PROGRAMS.ResourceID"

q_office_counts = "SELECT DISTINCT DisplayName0, COUNT(*) AS QTY FROM (SELECT DISTINCT DisplayName0, ResourceID " & _
	"FROM dbo.v_GS_ADD_REMOVE_PROGRAMS WHERE (DisplayName0 LIKE 'Microsoft Office Prof%') OR " & _
	"(DisplayName0 LIKE 'Microsoft Office 9%') OR (DisplayName0 LIKE 'Microsoft Office 365%') ) AS T1 " & _
	"GROUP BY DisplayName0"

q_devices = "SELECT DISTINCT dbo.v_R_System.ResourceID, dbo.v_R_System.Name0, dbo.v_R_System.AD_Site_Name0, dbo.v_R_System.Client0, " & _
	"dbo.v_R_System.Operating_System_Name_and0, dbo.v_GS_COMPUTER_SYSTEM.Domain0, dbo.v_GS_COMPUTER_SYSTEM.Manufacturer0, " & _
	"dbo.v_GS_COMPUTER_SYSTEM.Model0, dbo.v_GS_COMPUTER_SYSTEM.SystemType0, dbo.v_GS_COMPUTER_SYSTEM.PrimaryOwnerName0, " & _
	"dbo.v_GS_X86_PC_MEMORY.TotalPhysicalMemory0, dbo.v_GS_SYSTEM_ENCLOSURE.ChassisTypes0, dbo.v_GS_SYSTEM_ENCLOSURE.SerialNumber0, " & _
	"dbo.v_GS_OPERATING_SYSTEM.Caption0, dbo.v_GS_OPERATING_SYSTEM.CSDVersion0, dbo.v_GS_OPERATING_SYSTEM.CurrentTimeZone0, " & _
	"dbo.v_R_System.Full_Domain_Name0, dbo.v_GS_COMPUTER_SYSTEM.UserName0, dbo.v_R_System.Distinguished_Name0, " & _
	"dbo.v_R_System.Virtual_Machine_Host_Name0, dbo.v_R_System.Creation_Date0, dbo.v_R_System.CPUType0, " & _
	"dbo.v_R_System.Client_Version0 " & _
	"FROM dbo.v_R_System LEFT OUTER JOIN " & _
	"dbo.v_GS_COMPUTER_SYSTEM ON dbo.v_R_System.Name0 = dbo.v_GS_COMPUTER_SYSTEM.Name0 LEFT OUTER JOIN " & _
	"dbo.v_GS_OPERATING_SYSTEM ON dbo.v_R_System.ResourceID = dbo.v_GS_OPERATING_SYSTEM.ResourceID LEFT OUTER JOIN " & _
	"dbo.v_GS_SYSTEM_ENCLOSURE ON dbo.v_R_System.ResourceID = dbo.v_GS_SYSTEM_ENCLOSURE.ResourceID LEFT OUTER JOIN " & _
	"dbo.v_GS_X86_PC_MEMORY ON dbo.v_R_System.ResourceID = dbo.v_GS_X86_PC_MEMORY.ResourceID"

q_client_summary = "SELECT DISTINCT AD_Site_Name0, COUNT(*) AS Clients " & _
	"FROM (" & q_devices & ") AS T1 WHERE (T1.Client0 = 1) GROUP BY AD_Site_Name0"

q_client_sum = "SELECT T1.ADSiteName, COMPUTERS, CLIENTS, COMPUTERS-CLIENTS AS MISSING " & _
	"FROM (" & _
	"SELECT DISTINCT AD_Site_Name0 AS ADSiteName, COUNT(DISTINCT ResourceID) AS COMPUTERS " & _
	"FROM dbo.v_R_System " & _
	"WHERE (AD_Site_Name0 IS NOT NULL) " & _
	"GROUP BY AD_Site_Name0 " & _
	") AS T1 INNER JOIN (" & _
	"SELECT DISTINCT AD_Site_Name0 AS ADSiteName, COUNT(DISTINCT ResourceID) AS Clients " & _
	"FROM dbo.v_R_System " & _
	"WHERE (AD_Site_Name0 IS NOT NULL) AND (Client0 = 1) " & _
	"GROUP BY AD_Site_Name0 " & _
	") AS T2 ON T1.ADSiteName = T2.ADSiteName "

q_count_devices = "SELECT COUNT(DISTINCT ResourceID) AS QTY FROM dbo.v_R_System WHERE (Name0 LIKE '%VVV%')"
q_count_apps   = "SELECT COUNT(DISTINCT DisplayName0) AS QTY FROM dbo.v_GS_ADD_REMOVE_PROGRAMS WHERE (DisplayName0 LIKE '%VVV%')"
q_count_sites  = "SELECT COUNT(DISTINCT AD_Site_Name0) AS QTY FROM dbo.v_R_System WHERE (AD_Site_Name0 LIKE '%VVV%')"
q_count_models = "SELECT COUNT(DISTINCT Model0) AS QTY FROM dbo.v_GS_COMPUTER_SYSTEM WHERE (Model0 LIKE '%VVV%')"
q_count_users  = "SELECT COUNT(DISTINCT UserName0) AS QTY FROM dbo.v_GS_COMPUTER_SYSTEM WHERE (UserName0 LIKE '%VVV%')"
q_count_groups = "SELECT COUNT(DISTINCT Usergroup_Name0) AS QTY FROM dbo.v_R_UserGroup WHERE (Usergroup_Name0 LIKE '%VVV%')"
q_count_subnets = "SELECT COUNT(DISTINCT dbo.v_R_System.Name0) AS QTY FROM dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION INNER JOIN " & _
	"dbo.v_R_System ON dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.ResourceID = dbo.v_R_System.ResourceID WHERE (IPAddress0 LIKE '%VVV%')"
q_count_collections = "SELECT COUNT(DISTINCT Name) AS QTY FROM dbo.v_Collection WHERE (Name LIKE '%VVV%')"

q_search_devices = "SELECT DISTINCT [Name0] AS ItemName FROM dbo.v_R_System WHERE (Name0 LIKE '%VVV%')"
q_search_apps   = "SELECT DISTINCT DisplayName0 AS ItemName FROM dbo.v_GS_ADD_REMOVE_PROGRAMS WHERE (DisplayName0 LIKE '%VVV%')"
q_search_sites  = "SELECT DISTINCT AD_Site_Name0 AS ItemName FROM dbo.v_R_System WHERE (AD_Site_Name0 LIKE '%VVV%')"
q_search_models = "SELECT DISTINCT Model0 AS ItemName FROM dbo.v_GS_COMPUTER_SYSTEM WHERE (Model0 LIKE '%VVV%')"
q_search_users  = "SELECT DISTINCT User_Name0 AS ItemName FROM dbo.v_R_User WHERE (Full_User_Name0 LIKE '%VVV%') OR (User_Name0 LIKE '%VVV%')"
q_search_groups = "SELECT DISTINCT UserGroup_Name0 AS ItemName FROM dbo.v_R_UserGroup WHERE (UserGroup_Name0 LIKE '%VVV%')"
q_search_subnets = "SELECT DISTINCT dbo.v_R_System.Name0 AS ItemName, IPAddress0 AS IPAddress, AD_Site_Name0 AS ADSiteName " & _
	"FROM dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION INNER JOIN " & _
	"dbo.v_R_System ON dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION.ResourceID = dbo.v_R_System.ResourceID WHERE (IPAddress0 LIKE '%VVV%')"
q_search_collections = "SELECT DISTINCT [Name] AS ItemName FROM dbo.v_Collection WHERE (Name LIKE '%VVV%')"
%>