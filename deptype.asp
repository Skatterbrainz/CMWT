<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... deptype.asp
' lastupdate..... 12/04/2016
' description.... application deployment types
'-----------------------------------------------------------------------------
time1 = Timer
RecID   = CMWT_GET("id", "")
SortBy  = CMWT_GET("s", "AssignmentName")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE RecID, "Record name was not provided"

PageTitle = "Deployment Type"
PageBackLink = "deployments.asp"
PageBackName = "Deployments"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
query = "SELECT DISTINCT " & _
	"dbo.v_ApplicationAssignment.AssignmentID, " & _
	"dbo.v_ApplicationAssignment.ApplicationName, " & _
	"dbo.v_ApplicationAssignment.AssignmentName, " & _
	"dbo.v_ApplicationAssignment.Description, " & _
	"dbo.v_ApplicationAssignment.CollectionID, " & _
	"dbo.v_ApplicationAssignment.CollectionName, " & _
	"dbo.vAppDeploymentTargetingInfo.Technology, " & _
	"dbo.vAppDeploymentTargetingInfo.CollectionType, " & _
	"dbo.v_ApplicationAssignment.OfferTypeID, " & _
	"dbo.v_ApplicationAssignment.WoLEnabled, " & _
	"dbo.v_ApplicationAssignment.EnforcementEnabled, " & _
	"dbo.v_ApplicationAssignment.UserUIExperience, " & _
	"dbo.v_ApplicationAssignment.AssignmentEnabled, " & _
	"dbo.v_ApplicationAssignment.PersistOnWriteFilterDevices, " & _
	"dbo.v_ApplicationAssignment.RebootOutsideOfServiceWindows, " & _
	"dbo.v_ApplicationAssignment.OverrideServiceWindows, " & _
	"dbo.v_ApplicationAssignment.EvaluationSchedule, " & _
	"dbo.v_ApplicationAssignment.SuppressReboot, " & _
	"dbo.v_ApplicationAssignment.NotifyUser, " & _
	"dbo.v_ApplicationAssignment.LogComplianceToWinEvent, " & _
	"dbo.v_ApplicationAssignment.IncludeSubCollections, " & _
	"dbo.v_ApplicationAssignment.UseGMTTimes, " & _
	"dbo.v_ApplicationAssignment.CreationTime, " & _
	"dbo.v_ApplicationAssignment.ExpirationTime, " & _
	"dbo.v_ApplicationAssignment.StartTime, " & _
	"dbo.v_ApplicationAssignment.EnforcementDeadline, " & _
	"dbo.v_ApplicationAssignment.SoftDeadlineEnabled, " & _
	"dbo.v_ApplicationAssignment.LastModificationTime, " & _
	"dbo.v_ApplicationAssignment.LastModifiedBy, " & _
	"dbo.v_ApplicationAssignment.UpdateDeadlineTime, " & _
	"dbo.v_ApplicationAssignment.Priority " & _
	"FROM dbo.v_ApplicationAssignment " & _
	"INNER JOIN " & _
	"dbo.vAppDeploymentTargetingInfo ON " & _
	"dbo.v_ApplicationAssignment.AssignmentID = dbo.vAppDeploymentTargetingInfo.AssignmentID " & _
	"INNER JOIN " & _
	"dbo.vDeploymentSummary ON " & _
	"dbo.v_ApplicationAssignment.AssignmentID = dbo.vDeploymentSummary.AssignmentID " & _
	"WHERE " & _
	"dbo.v_ApplicationAssignment.AssignmentID='" & RecID & "' " & _
	"ORDER BY " & SortBy
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query

Do Until rs.EOF
	Response.Write "<table class=""tfx"">"
	For i = 0 to rs.Fields.Count -1
		fn = rs.Fields(i).Name
		fv = rs.Fields(i).Value
		fv = CMWT_AutoLink (fn, fv)
		Response.Write "<tr class=""tr1""><td class=""td6 v10 w150 bgDarkGray"">" & fn & "</td>" & _
			"<td class=""td6 v10"">" & fv & "</td></tr>"
	Next
	Response.Write "</table><br/>"
	rs.MoveNext
Loop
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>