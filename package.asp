<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... package.asp
' lastupdate..... 12/07/2016
' description.... package details report
'-----------------------------------------------------------------------------
time1 = Timer
RecID   = CMWT_GET("id", "")
KeySet  = CMWT_GET("ks", "1")
SortBy  = CMWT_GET("s", "DPServer")
KeySet2 = CMWT_GET("k2", "0")
QueryOn = CMWT_GET("qq", "")
CMWT_VALIDATE RecID, "Record name was not provided"

query = "SELECT TOP 1 " & _
	"Name,PackageID,Version,Language,Manufacturer AS Publisher,Description," & _
	"PkgSourcePath,StoredPkgPath,SourceVersion,SourceDate," & _
	"ShareType,ShareName,SourceSite,ForcedDisconnectEnabled,ForcedDisconnectNumRetries," & _
	"ForcedDisconnectDelay,Priority,PreferredAddressType,IgnoreAddressSchedule," & _
	"LastRefreshTime,MIFFilename,MIFPublisher,MIFName," & _
	"MIFVersion,ActionInProgress,ImageFlags,PackageType,SecurityKey " & _
	"FROM dbo.v_Package " & _
	"WHERE PackageID = '" & RecID & "'"

CMWT_DEBUG query

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
PkgName = rs.Fields("Name").value
PageTitle = PkgName
If KeySet2 = "8" Then
	PageBackLink = "applications.asp"
	PageBackName = "Applications"
Else
	PageBackLink = "packages.asp"
	PageBackName = "Packages"
End If

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%	
menulist = "1=General,2=Distribution,3=Programs,4=Notes"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	mlink = "package.asp?id=" & RecID & "&ks=" & mset(0)
	If KeySet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

Select Case KeySet
	Case "1"
		CMWT_DB_TABLEROWGRID rs, "", "", ""
		CMWT_DB_CLOSE()
	Case "2"
		CMWT_DB_CLOSE()
		query = "SELECT dbo.v_DistributionPoints.ServerName AS DPServer, " & _
			"dbo.v_DistributionPoint.PackageID, " & _
			"dbo.v_ContentDistributionReport.StateName, " & _
			"dbo.v_ContentDistributionReport.SummaryDate " & _
			"FROM dbo.v_DistributionPoint INNER JOIN " & _
			"dbo.v_DistributionPoints ON " & _
			"dbo.v_DistributionPoint.ServerNALPath = dbo.v_DistributionPoints.NALPath " & _
			"INNER JOIN dbo.v_ContentDistributionReport ON " & _
			"dbo.v_DistributionPoint.PackageID = dbo.v_ContentDistributionReport.PkgID " & _
			"WHERE dbo.v_DistributionPoint.PackageID='" & RecID & "' " & _
			"ORDER BY " & SortBY
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEGRID rs, "", "package.asp?ks=2", ""
		CMWT_DB_CLOSE()
	Case "3"
		CMWT_DB_CLOSE()
		query = "SELECT DISTINCT ProgramName,CommandLine," & _
			"Comment,Description,Requirements,DependentProgram," & _
			"DriveLetter AS DrvLetter,WorkingDirectory," & _
			"DiskSpaceRequired AS DiskSpace,Duration,RemovalKey " & _
			"FROM dbo.v_Program " & _
			"WHERE PackageID='" & RecID & "' " & _
			"ORDER BY ProgramName"
		CMWT_DB_QUERY Application("DSN_CMDB"), query
		CMWT_DB_TABLEGRID rs, "", "package.asp?ks=3", ""
		CMWT_DB_CLOSE()
	Case "4"
		query = "SELECT NoteID, Comment, DateCreated, CreatedBy " & _
		"FROM dbo.Notes " & _
		"WHERE (AttachedTo = '" & RecID & "') AND (AttachClass = 'PACKAGE') " & _
		"ORDER BY NoteID DESC"
	
		Response.Write "<table class=""tfx"">"
		
		CMWT_DB_QUERY Application("DSN_CMWT"), query

		If Not(rs.BOF And rs.EOF) Then
			found = True
			xrows = rs.RecordCount
			xcols = rs.Fields.Count

			Response.Write "<tr>"
			For i = 0 to xcols-1
				Response.Write "<td class=""td6 v10 bgGray"">" & rs.Fields(i).Name & "</td>"
			Next
			Response.Write "</tr>"

			Do Until rs.EOF
				Response.Write "<tr class=""tr1"">"
				For i = 0 to xcols-1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					Select Case Ucase(fn)
						Case "NOTEID":
							fv = CMWT_IMG_LINK (TRUE, "icon_del2", "icon_del1", "icon_del3", "confirm.asp?id=" & fv & "&tn=notes&pk=noteid&t=package.asp|id=" & RecID & "^set=10", "Remove") & " " & _
								CMWT_IMG_LINK (TRUE, "icon_edit2", "icon_edit1", "icon_edit2", "noteedit.asp?id=" & fv, "Edit")
							Response.Write "<td class=""td6 v10 w50"">" & fv & "</td>"
						Case Else:
							Response.Write "<td class=""td6 v10"">" & fv & "</td>"
					End Select
					
				Next
				Response.Write "</tr>"
				rs.MoveNext
			Loop
			Response.Write "<tr>" & _
				"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
				xrows & " rows returned</td></tr>"
		Else
			Response.Write "<tr class=""h100 tr1"">" & _
				"<td class=""td6 v10 ctr"">No matching rows returned</td></tr>"
		End If
		
		Response.Write "</table>"
		
		If CMWT_ADMIN() Then
			Response.Write "<br/><table class=""tfx""><tr><td class=""v10"">" & _
				"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" " & _
				"value=""New Note"" onClick=""document.location.href='noteadd.asp?id=" & RecID & "&t=package'"" " & _
				"title=""New Note"" /></td></tr></table>"
		End If

End Select

CMWT_SHOW_QUERY()
CMWT_Footer()
%>

</body>
</html>