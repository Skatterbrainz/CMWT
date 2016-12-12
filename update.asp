<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... update.asp
' lastupdate..... 12/08/2016
' description.... deployment status summary
'-----------------------------------------------------------------------------
time1 = Timer
UpdateID = CMWT_GET("id", "")
CMWT_VALIDATE UpdateID, "Update CID was not provided"

QueryOn   = CMWT_GET("qq", "")
PageTitle = "Software Update"
PageBackLink = "updates.asp"
PageBackName = "Software Updates"

subselect = ""

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT TOP 1 " & _
	"dbo.v_UpdateInfo.Title, dbo.v_UpdateInfo.Description, " & _
	"dbo.v_UpdateInfo.InfoURL, dbo.v_UpdateInfo.UpdateType, dbo.v_UpdateInfo.BulletinID, " & _
	"dbo.v_UpdateInfo.ArticleID, dbo.vSMS_SoftwareUpdate.SeverityName, " & _
	"dbo.vSMS_SoftwareUpdate.NumTotal AS Scanned, dbo.vSMS_SoftwareUpdate.NumMissing AS Missing, " & _
	"dbo.vSMS_SoftwareUpdate.NumPresent AS Installed, " & _
	"dbo.vSMS_SoftwareUpdate.NumNotApplicable AS NotRequired, dbo.vSMS_SoftwareUpdate.NumUnknown AS Unknown, " & _
	"dbo.vSMS_SoftwareUpdate.PercentCompliant AS Compliant,dbo.vSMS_SoftwareUpdate.LastStatusTime, " & _
	"dbo.v_UpdateInfo.ModelId, dbo.v_UpdateInfo.DateCreated, dbo.v_UpdateInfo.DateLastModified, " & _
	"dbo.v_UpdateInfo.ModelName, dbo.v_UpdateInfo.LastModifiedBy, dbo.v_UpdateInfo.CreatedBy, " & _
	"dbo.v_UpdateInfo.PermittedUses, dbo.v_UpdateInfo.IsBundle, dbo.v_UpdateInfo.IsHidden, " & _
	"dbo.v_UpdateInfo.IsTombstoned, dbo.v_UpdateInfo.IsUserDefined, dbo.v_UpdateInfo.IsEnabled, " & _
	"dbo.v_UpdateInfo.IsExpired, dbo.v_UpdateInfo.SourceSite, dbo.v_UpdateInfo.ContentSourcePath, " & _
	"dbo.v_UpdateInfo.ApplicabilityCondition, dbo.v_UpdateInfo.Precedence, dbo.v_UpdateInfo.EULAExists, " & _
	"dbo.v_UpdateInfo.EULAAccepted, dbo.v_UpdateInfo.EULASignoffDate, dbo.v_UpdateInfo.EULASignoffUser, " & _
	"dbo.v_UpdateInfo.IsQuarantined, dbo.v_UpdateInfo.ModifiedTime, dbo.v_UpdateInfo.IsDeployed, " & _
	"dbo.v_UpdateInfo.IsSuperseded, dbo.v_UpdateInfo.IsChild, dbo.v_UpdateInfo.InUse, " & _
	"dbo.v_UpdateInfo.IsLatest, dbo.v_UpdateInfo.IsBroken, dbo.v_UpdateInfo.ConfigurationFlags, " & _
	"dbo.v_UpdateInfo.EffectiveDate, dbo.v_UpdateInfo.PlatformType, dbo.v_UpdateInfo.IsUserCI, " & _
	"dbo.v_UpdateInfo.SDMPackageVersion, dbo.v_UpdateInfo.SDMPackageDigest, dbo.v_UpdateInfo.CI_UniqueID, dbo.v_UpdateInfo.CIType_ID, " & _
	"dbo.v_UpdateInfo.CI_CRC, dbo.v_UpdateInfo.CIVersion, dbo.v_UpdateInfo.RevisionTag, " & _
	"dbo.v_UpdateInfo.IsSignificantRevision, dbo.v_UpdateInfo.SedoObjectVersion, dbo.v_UpdateInfo.ApplicableAtUserLogon, " & _
	"dbo.v_UpdateInfo.CustomSeverity, dbo.v_UpdateInfo.DatePosted, dbo.v_UpdateInfo.DateRevised, " & _
	"dbo.v_UpdateInfo.RevisionNumber, dbo.v_UpdateInfo.MaxExecutionTime, " & _
	"dbo.v_UpdateInfo.RequiresExclusiveHandling, dbo.v_UpdateInfo.UpdateSource_ID, " & _
	"dbo.v_UpdateInfo.MinSourceVersion, dbo.v_UpdateInfo.LocaleID, dbo.v_UpdateInfo.Locales " & _
	"FROM dbo.v_UpdateInfo INNER JOIN " & _
		"dbo.vSMS_SoftwareUpdate ON dbo.v_UpdateInfo.CI_ID = dbo.vSMS_SoftwareUpdate.CI_ID " & _
	"WHERE (dbo.v_UpdateInfo.CI_ID = '" & UpdateID & "')"

'query = "SELECT TOP 1 " & _
'	"Title,Description,InfoURL,UpdateType,BulletinID,ArticleID,Severity," & _
'	"ModelId,DateCreated,DateLastModified,ModelName,LastModifiedBy,CreatedBy," & _
'	"PermittedUses,IsBundle,IsHidden,IsTombstoned,IsUserDefined,IsEnabled,IsExpired," & _
'	"SourceSite,ContentSourcePath,ApplicabilityCondition,Precedence,EULAExists," & _
'	"EULAAccepted,EULASignoffDate,EULASignoffUser,IsQuarantined,ModifiedTime,IsDeployed," & _
'	"IsSuperseded,IsChild,InUse,IsLatest,IsBroken,ConfigurationFlags,EffectiveDate," & _
'	"PlatformType,IsUserCI,SDMPackageVersion,SDMPackageDigest,CI_UniqueID,CIType_ID," & _
'	"CI_CRC,CIVersion,RevisionTag,IsSignificantRevision,SedoObjectVersion,ApplicableAtUserLogon," & _
'	"CustomSeverity,DatePosted,DateRevised,RevisionNumber,MaxExecutionTime,RequiresExclusiveHandling," & _
'	"UpdateSource_ID,MinSourceVersion,LocaleID,Locales " & _
'	"FROM dbo.v_UpdateInfo " & _
'	"WHERE CI_ID='" & UpdateID & "'"

CMWT_DEBUG query

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TableRowGrid rs, "", "", ""
'CMWT_DB_TableGridFilter rs, "", "updates.asp", "", colset, "updates.asp?fn=X&fv=Y"
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_FOOTER()

Response.Write "</body></html>"
%>
