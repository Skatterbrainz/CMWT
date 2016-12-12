<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... collmem.asp
' lastupdate..... 12/08/2016
' description.... collection direct-rule membership tools
'-----------------------------------------------------------------------------
time1 = Timer

CollectionID1 = CMWT_GET("cid1", "")
CollectionID2 = CMWT_GET("cid2", "")
ActionType    = CMWT_GET("xx3", "")

CMWT_VALIDATE CollectionID1, "Source Collection ID was not specified"

'-----------------------------------------------------------------------------
' sub-name: CMWT_CM_ListCollectionMembers
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_CM_ListCollectionMembers (c, CollID)
	Dim query, cmd, rs, x1, x2
	If CollID <> "" Then
		query = "SELECT DISTINCT Name0, ResourceID " &_
			"FROM dbo.v_R_System " &_
			"WHERE dbo.v_R_System.ResourceID IN " &_
			"(SELECT DISTINCT DBO.V_FULLCOLLECTIONMEMBERSHIP.ResourceID " &_
			"FROM DBO.V_FULLCOLLECTIONMEMBERSHIP " &_
			"WHERE (DBO.V_FULLCOLLECTIONMEMBERSHIP.COLLECTIONID = '" & CollID & "')) " &_
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
	Else
		Response.Write "<option value=""""></option>"
	End If
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_CM_ListDirectCollections
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_CM_ListDirectCollections (c, DefaultID, IdNum, ExcludeID)
	Dim query, cmd, rs, x1, x2, AltName
	query = "SELECT DISTINCT Name, CollectionID " & _
		"FROM dbo.v_Collection " & _
		"WHERE (CollectionID NOT IN " & _
        "(SELECT DISTINCT CollectionID " & _
        "FROM dbo.v_CollectionRuleQuery AS v_CollectionRuleQuery_1)) AND (CollectionType = 2) " & _
		"ORDER BY Name"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If IdNum = 1 Then
		TagName = "cid1="
		AltName = "cid2=" & CollectionID2
	Else
		TagName = "cid2="
		AltName = "cid1=" & CollectionID1
	End If
	Do Until rs.EOF
		x1 = rs.Fields("CollectionID").value
		x2 = rs.Fields("Name").value
		If Ucase(x1) = Ucase(DefaultID) Then
			Response.Write "<option value=""collmem.asp?" & TagName & x1 & "&" & AltName & """ selected>" & x2 & "</option>"
		Else
			Response.Write "<option value=""collmem.asp?" & TagName & x1 & "&" & AltName & """>" & x2 & "</option>"
		End If
		rs.MoveNext
	Loop
	rs.Close
End Sub

PageTitle    = "Collection Tools"
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Dim conn, cmd, rs, query
CMWT_DB_OPEN Application("DSN_CMDB")
CollName = CMWT_CM_ObjectProperty (conn, "v_Collections", "SiteID", KeyValue, "CollectionName")
CollType = CMWT_CM_ObjectProperty (conn, "v_Collections", "SiteID", KeyValue, "CollectionType")
CMWT_DB_CLOSE()

CMWT_DB_OPEN Application("DSN_CMDB")

'if (CMWT_CM_CollectionRuleType( conn, KeyValue )) = "DIRECT" Then
%>
<form name="form1" id="form1" method="post" action="collmem2.asp">
<table class="tfx">
	<tr>
		<td class="td6 v10 w420">
			<h2>Source Collection</h2>
			<select name="c1" id="c1" size="1" class="pad5 v10 w400" onChange="if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }">
				<option value=""></option>
				<%
				CMWT_CM_ListDirectCollections conn, CollectionID1, 1, ""
				%>
			</select>
		</td>
		<td class="td6 v10">
			<h2>Target Collection</h2>
			<select name="c2" id="c2" size="1" class="pad5 v10 w400" onChange="if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }">
				<option value=""></option>
				<%
				CMWT_CM_ListDirectCollections conn, CollectionID2, 2, ""
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td6 v10 w420">
			<select name="m1" id="m1" size="10" class="pad5 v10 w400" multiple=true>
				<%
				CMWT_CM_ListCollectionMembers conn, CollectionID1
				%>
			</select>
			<select name="a1" id="a1" size="1" class="pad5 v10 w400">
				<option></option>
				<option value="COPY">Copy to Target</option>
				<option value="MOVE">Move to Target</option>
			</select>
		</td>
		<td class="td6 v10">
			<select name="m2" id="m2" size="10" class="pad5 v10 w400" multiple=true>
				<%
				CMWT_CM_ListCollectionMembers conn, CollectionID2
				%>
			</select>
			<select name="a2" id="a2" size="1" class="pad5 v10 w400">
				<option></option>
				<option value="COPY">Copy to Source</option>
				<option value="MOVE">Move to Source</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="td6 v10" colspan="2">
			<input type="hidden" name="cid1" id="cid1" value="<%=CollectionID1%>" />
			<input type="hidden" name="cid2" id="cid2" value="<%=CollectionID2%>" />
			<input type="reset" name="b0" id="b0" class="btx w140 h30" value="Reset" title="Clear Selections" />
			<input type="submit" name="b1" id="b1" class="btx w140 h30" value="Execute!" title="Execute Now!" />
		</td>
	</tr>
</table>
</form>

<%
CMWT_DB_CLOSE()
CMWT_FOOTER()
Response.Write "</body></table>"
%>