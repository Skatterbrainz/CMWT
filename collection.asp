<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... collection.asp
' lastupdate..... 12/08/2016
' description.... collection details report
'-----------------------------------------------------------------------------
time1 = Timer
SortBy   = CMWT_GET("s", "CollectionName")
KeyValue = CMWT_GET("id", "")
AltValue = CMWT_GET("cn", "")
KeySet   = CMWT_GET("ks", "1")
QueryOn  = CMWT_GET("qq", "")

'-----------------------------------------------------------------------------
' function-name: DPMS_CM_RESOURCETYPENAME
' function-desc: 
'-----------------------------------------------------------------------------

Function DPMS_CM_RESOURCETYPENAME (n)
	Select Case n
		Case 1: DPMS_CM_RESOURCETYPENAME = "USER"
		Case 2: DPMS_CM_RESOURCETYPENAME = "DEVICE"
		Case Else: DPMS_CM_RESOURCETYPENAME = "?"
	End Select
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_CM_ListCollectionMembers
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_CM_ListCollectionMembers (c, ResourceType, DefaultResID)
	Dim query, cmd, rs, x1, x2
	query = "SELECT DISTINCT Name0, ResourceID " &_
		"FROM dbo.v_R_System " &_
		"WHERE Name0 NOT IN " &_
		"(SELECT DISTINCT DBO.V_FULLCOLLECTIONMEMBERSHIP.NAME AS MEMBER " &_
		"FROM DBO.V_FULLCOLLECTIONMEMBERSHIP " &_
		"WHERE (DBO.V_FULLCOLLECTIONMEMBERSHIP.COLLECTIONID = '" & KeyValue & "')) " &_
		"AND Client0=1 " &_
		"ORDER BY Name0"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	Do Until rs.EOF
		x1 = rs.Fields("ResourceID").value
		x2 = rs.Fields("Name0").value
		Response.Write "<option value=""" & x2 & """>" & x2 & "</option>"
		rs.MoveNext
	Loop
	rs.Close
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_CM_ListDirectCollections
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_CM_ListDirectCollections (c)
	Dim query, cmd, rs, x1, x2
	query = "SELECT DISTINCT dbo.v_CollectionRuleDirect.CollectionID, dbo.v_Collection.Name " & _
		"FROM dbo.v_CollectionRuleDirect INNER JOIN " & _
		"dbo.v_Collection ON dbo.v_CollectionRuleDirect.CollectionID = dbo.v_Collection.CollectionID " & _
		"ORDER BY dbo.v_Collection.Name"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	Do Until rs.EOF
		x1 = rs.Fields("CollectionID").value
		x2 = rs.Fields("Name").value
		Response.Write "<option value=""" & x1 & """>" & x2 & "</option>"
		rs.MoveNext
	Loop
	rs.Close
End Sub

Dim conn, cmd, rs, query
CMWT_DB_OPEN Application("DSN_CMDB")
CollName = CMWT_CM_ObjectProperty (conn, "v_Collections", "SiteID", KeyValue, "CollectionName")
CollType = CMWT_CM_ObjectProperty (conn, "v_Collections", "SiteID", KeyValue, "CollectionType")
CMWT_DB_CLOSE()

PageTitle = CollName
PageBackLink = "collections.asp?ks=" & CollType
PageBackName = "Collections"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "1=General,2=Members,3=Queries,4=Variables,5=Notes,6=Tools"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	mlink = "collection.asp?id=" & KeyValue & "&ks=" & mset(0)
	If KeySet = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

'----------------------------------------------------------------

Select Case KeySet
	Case "1":

		query = "SELECT TOP 1 " & _
			"CollectionName," & _
			"dbo.v_Collections.CollectionID, " & _
			"dbo.v_Collections.SiteID," & _
			"ResultTableName,CollectionComment,Schedule,SourceLocaleID," & _
			"LastChangeTime,LastRefreshRequest," & _
			"CASE WHEN CollectionType=2 THEN 'DEVICE' " & _
			"ELSE 'USER' END AS CollectionType," & _
			"LimitToCollectionID," & _
			"IsReferenceCollection,BeginDate,EvaluationStartTime,LastRefreshTime," & _
			"LastIncrementalRefreshTime,LastMemberChangeTime,CurrentStatus," & _
			"CurrentStatusTime,LimitToCollectionName,ISVData,ISVString,Flags," & _
			"CollectionVariablesCount,ServiceWindowsCount,PowerConfigsCount," & _
			"RefreshType,MonitoringFlags,IsBuiltIn,IncludeExcludeCollectionsCount," & _
			"MemberCount,LocalMemberCount,ResultClassName,HasProvisionedMember, " & _
			"CASE WHEN (dbo.v_CollectionRuleDirect.RuleName <> '') THEN 'DIRECT' ELSE 'QUERY' END AS RuleType " & _
			"FROM dbo.v_Collections LEFT OUTER JOIN " & _
			"dbo.v_CollectionRuleDirect ON dbo.v_Collections.SiteID = dbo.v_CollectionRuleDirect.CollectionID " & _
			"WHERE (SiteID='" & KeyValue & "')"

		CMWT_DB_OPEN Application("DSN_CMDB")
		Set cmd  = Server.CreateObject("ADODB.Command")
		Set rs   = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		Set cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd
		Response.Write "<table class=""tfx"">"
		Do Until rs.EOF
			For i = 0 to rs.Fields.Count - 1
				fn = rs.Fields(i).Name
				fv = Trim(rs.Fields(i).Value)
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 w180 v10 bgDarkGray"">" & fn & "</td>" & _
					"<td class=""td6 v10"">" & CMWT_AutoLink(fn,fv) & "</td></tr>"
			Next
			rs.MoveNext
		Loop
		Response.Write "</table>"
		CMWT_DB_CLOSE()
	
	Case "2":

		query = "SELECT DISTINCT " & _
			"dbo.v_FullCollectionMembership.Name AS MemberName, " & _
			"dbo.v_FullCollectionMembership.ResourceID, " & _
			"dbo.v_FullCollectionMembership.ResourceType, " & _
			"dbo.v_FullCollectionMembership.Domain, " & _
			"dbo.v_FullCollectionMembership.SMSID, " & _
			"dbo.v_FullCollectionMembership.SiteCode, " & _ 
			"dbo.v_Collection.Name AS CollectionName " & _
			"FROM dbo.v_FullCollectionMembership INNER JOIN " & _
			"dbo.v_Collection ON " & _
			"dbo.v_FullCollectionMembership.CollectionID = dbo.v_Collection.CollectionID " & _
			"WHERE (dbo.v_FullCollectionMembership.CollectionID = '" & KeyValue & "')"

		CMWT_DB_OPEN Application("DSN_CMDB")
		Set cmd  = Server.CreateObject("ADODB.Command")
		Set rs   = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		Set cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd
		If Not(rs.BOF And rs.EOF) Then
			found = True
		End If

		Response.Write "<form name=""form3"" id=""form3"" method=""post"" action=""cmcx.asp"">" & _
			"<input type=""hidden"" name=""cid"" id=""cid"" value=""" & KeyValue & """ />"
		
		If Not(rs.BOF and rs.EOF) Then
			xrows = rs.RecordCount
			Response.Write "<table class=""tfx""><tr><td class=""td6 w40 ctr bgGray""></td>"
			For i = 0 to rs.Fields.Count - 1
				fn = rs.Fields(i).Name
				If Lcase(fn) <> "collectionname" Then
					Response.Write "<td class=""td6 v10 bgGray"">" & fn & "</td>"
				End If
			Next 
			Response.Write "</tr>"
			Do Until rs.EOF
				Response.Write "<tr class=""tr1"">"
				rtype = rs.Fields("ResourceType").Value
				For i = 0 to rs.Fields.Count - 1
					fn = rs.Fields(i).Name
					fv = rs.Fields(i).Value
					If (fn = "MemberName") Then
						Response.Write "<td class=""td6 w40 ctr"">" & _
							CMWT_IMG_LINK (CMWT_ADMIN(), "icon_del2", "icon_del1", "icon_del2", "cmcx.asp?cid=" & KeyValue & "&mx=rem&cn=" & fv, "Remove from Collection") & _
							"</td>"
						if rtype = 4 then
							fv = "<a href=""user.asp?cn=" & fv & """ title=""User Information"">" & fv & "</a>"
						else
							fv = "<a href=""device.asp?cn=" & fv & """ title=""Device Information"">" & fv & "</a>"
						end if
					End If
					If Lcase(fn) <> "collectionname" Then
						Response.Write "<td class=""td6 v10"">" & fv & "</td>"
					End If
				Next 
				Response.Write "</tr>"
				rs.MoveNext
			Loop
			Response.Write "<tr><td class=""td6 bgGray v10"" colspan=""7"">" & _
				xrows & " members found</td></tr></table>"
		Else
			Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
				"<td class=""td6 v10 ctr"">No members were found</td></tr></table>"
		End If

		If CMWT_ADMIN() Then
			If CollType = 2 Then
				Response.Write "<table class=""tfx"">" & _
					"<tr>" & _
					"<td class=""pad6 v10"">" & _
					"<form name=""form3"" id=""form3"" method=""post"" action=""cmcx.asp"">" & _
					"<input type=""hidden"" name=""mx"" id=""mx"" value=""ADD"" />" & _
					"<table class=""tfx""><tr><td class=""pad6 v10"">" & _
					"<select name=""cn"" id=""cn"" size=""1"" class=""w400 pad6"" title=""Select Device to Add..."">" & _
						"<option value=""""></option>"

				CMWT_CM_ListCollectionMembers conn, 2, ""

				Response.Write "</select> " & _
					"<input type=""submit"" name=""bx1"" id=""bx1"" class=""w140 h32 btx"" value=""Add"" />" & _
					"</form></td></tr></table>"
			End If
		End If
	
		CMWT_DB_CLOSE()
		
	Case "3":
			
		Response.Write "<table class=""tfx"">"

		query = "SELECT DISTINCT RuleName, QueryID, QueryExpression " & _
			"FROM CM_" & Application("CM_SITECODE") & ".dbo.v_CollectionRuleQuery " & _
			"WHERE CollectionID = '" & KeyValue & "' " & _
			"ORDER BY QueryID"
	
		CMWT_DB_OPEN Application("DSN_CMDB")
		Set cmd  = Server.CreateObject("ADODB.Command")
		Set rs   = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		Set cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd

		If Not(rs.BOF And rs.EOF) Then
			Response.Write "<tr>" & _
				"<td class=""td6 v10 bgGray"">ID</td>" & _
				"<td class=""td6 v10 bgGray"">Rule</td>" & _
				"<td class=""td6 v10 bgGray"">Query Expression</td></tr>"
			Do Until rs.EOF
				x1 = rs.Fields("QueryID").value
				x2 = rs.Fields("RuleName").value
				x3 = rs.Fields("QueryExpression").value
				
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10"">" & x1 & "</td>" & _
					"<td class=""td6 v10"">" & x2 & "</td>" & _
					"<td class=""td6 v10"">" & CMWT_PrettySQL(x3) & "</td>" & _
					"</tr>"
				rs.MoveNext
			Loop
		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No query rules were found</td></tr>"
		End If

		CMWT_DB_CLOSE()
		Response.Write "</table>"
	
	Case "4":
	
		query = "SELECT DISTINCT Name,Value,CASE WHEN IsMasked=1 THEN 'YES' ELSE 'NO' END AS Masked FROM dbo.v_CollectionVariable " & _
			"WHERE CollectionID='" & KeyValue & "' ORDER BY Name"

		CMWT_DB_OPEN Application("DSN_CMDB")
		Set cmd  = Server.CreateObject("ADODB.Command")
		Set rs   = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		Set cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd

		Response.Write "<table class=""tfx"">"
		
		If Not(rs.BOF And rs.EOF) Then
			Response.Write "<tr>" & _
				"<td class=""td6 v10 bgGray"">Name</td>" & _
				"<td class=""td6 v10 bgGray"">Value</td>" & _
				"<td class=""td6 v10 bgGray"">Masked</td></tr>"
			x22 = ""
			Do Until rs.EOF
				x1 = rs.Fields("Name").value
				x2 = rs.Fields("Value").value
				x3 = rs.Fields("Masked").value
				tmp = x2
				Do While tmp <> ""
					x22 = x22 & Left(tmp,72) & "<br/>"
					tmp = Mid(tmp,73)
				Loop
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10"">" & x1 & "</td>" & _
					"<td class=""td6 v10"">" & x22 & "</td>" & _
					"<td class=""td6 v10"">" & x3 & "</td>" & _
					"</tr>"
				rs.MoveNext
			Loop
		Else
			Response.Write "<tr class=""h100 tr1""><td class=""td6 v10 ctr"">No assigned collection variables were found</td></tr>"
		End If

		CMWT_DB_CLOSE()
		Response.Write "</table>"			

	Case "5":
		
		query = "SELECT NoteID, Comment, DateCreated, CreatedBy " & _
			"FROM dbo.Notes " & _
			"WHERE (AttachedTo = '" & KeyValue & "') AND (AttachClass = 'COLLECTION') " & _
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
							fv = CMWT_IMG_LINK (TRUE, "icon_del2", "icon_del1", "icon_del3", "confirm.asp?id=" & fv & "&tn=notes&pk=noteid&t=collection.asp^id=" & KeyValue & "^set=10^ks=4", "Remove") & " " & _
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
				xrows & " items returned</td></tr>"
		Else
			Response.Write "<tr class=""h100 tr1"">" & _
				"<td class=""td6 v10 ctr"">No assigned notes were found</td></tr>"
		End If
		
		Response.Write "</table>"
		
		If CMWT_ADMIN() Then
			Response.Write "<br/><table class=""tfx""><tr><td class=""v10"">" & _
				"<input type=""button"" name=""b1"" id=""b1"" class=""btx w150 h32"" " & _
				"value=""New Note"" onClick=""document.location.href='noteadd.asp?id=" & KeyValue & "&t=collection'"" " & _
				"title=""New Note"" /></td></tr></table>"
		End If
		CMWT_DB_CLOSE()
	
	Case "6":
		
		CMWT_Hide_QueryLink = True
		Response.Write "<table class=""tfx"">"
		
		CMWT_DB_OPEN Application("DSN_CMDB")

		if (CMWT_CM_CollectionRuleType( conn, KeyValue )) = "DIRECT" Then
			%>
			<tr class="bgDarkGray vtop">
				<td class="td6a v10 w250">
					<p>Invoke Client Action</p>
					<form name="formx1" id="formx1" method="post" action="">
						<select name="xx1" id="xx1" size="6" class="pad5 v10 w200">
							<option>Client Machine Policy Refresh</option>
							<option>Client Discovery Cycle</option>
							<option>Hardware Inventory Cycle</option>
							<option>Software Inventory Cycle</option>
							<option>Software Updates Scan Cycle</option>
						</select>
						<p><input type="submit" name="bx1" id="bx1" class="btx w200 h30" value="Execute" title="Execute" /></p>
					</form>
				</td>
				<td class="td6a v10 w250">
					<p>Execute Tools</p>
					<form name="formx2" id="formx2" method="post" action="">
						<select name="xx2" id="xx2" size="6" class="pad5 v10 w200">
							<option>Restart Members</option>
							<option>Shutdown Members</option>
							<option>Group Policy Update</option>
							<option>Restart SMSAgent Service</option>
							<option>Ping All</option>
						</select>
						<p><input type="submit" name="bx2" id="bx2" class="btx w200 h30" value="Execute" title="Execute" /></p>
					</form>
				</td>
				<td class="td6a v10">
					<p>Collection Members</p>
					<form name="formx3" id="formx3" method="post" action="">
						<select name="xx3" id="xx3" size="1" class="pad5 v10 w400">
							<option value=""></optioN>
						</select>
						<p><select name="xx3" id="xx3" size="1" class="pad5 v10 w200">
							<option value=""></option>
							<option value="COMPARE">Compare Members</option>
							<option value="COPY">Copy Members</option>
							<option value="MOVE">Move Members</option>
						</select></p>
						<p><input type="submit" name="bx3" id="bx3" class="btx w200 h30" value="Execute" title="Execute" /></p>
					</form>
				</td>
			</tr>
			<%
		else
			%>
			<tr class="bgDarkGray vtop">
				<td class="td6a v10 w250">
					<p>Invoke Client Action</p>
					<form name="formx1" id="formx1" method="post" action="">
						<select name="xx1" id="xx1" size="6" class="pad5 v10 w200">
							<option>Client Machine Policy Refresh</option>
							<option>Client Discovery Cycle</option>
							<option>Hardware Inventory Cycle</option>
							<option>Software Inventory Cycle</option>
							<option>Software Updates Scan Cycle</option>
						</select>
						<p><input type="submit" name="bx1" id="bx1" class="btx w200 h30" value="Execute" title="Execute" /></p>
					</form>
				</td>
				<td class="td6a v10 w250">
					<p>Execute Tools</p>
					<form name="formx2" id="formx2" method="post" action="">
						<select name="xx2" id="xx2" size="6" class="pad5 v10 w200">
							<option>Restart Members</option>
							<option>Shutdown Members</option>
							<option>Group Policy Update</option>
							<option>Restart SMSAgent Service</option>
							<option>Ping All</option>
						</select>
						<p><input type="submit" name="bx2" id="bx2" class="btx w200 h30" value="Execute" title="Execute" /></p>
					</form>
				</td>
				<td class="td6a v10">
					<p>Compare Collection Members</p>
					<form name="formx3" id="formx3" method="post" action="">
						<select name="xx3" id="xx3" size="1" class="pad5 v10 w400">
							<option value=""></option>
							<%
							CMWT_CM_ListDirectCollections conn
							%>
						</select>
						<input type="submit" name="bx3" id="bx3" class="btx w30 h30" value="..." title="Process" />
					</form>
				</td>
			</tr>
			<%
		end if
		CMWT_DB_CLOSE()
		
		Response.Write "</table>"
		
End Select

CMWT_SHOW_QUERY()
CMWT_Footer()	
%>

</body>
</html>