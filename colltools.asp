<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... colltools.asp
' lastupdate..... 11/30/2016
' description.... collection tools
'-----------------------------------------------------------------------------
time1 = Timer

QueryOn   = CMWT_GET("qq", "")
PageTitle = "Collection Tools"

Sub CMWT_Collections ()
	Dim query, conn, cmd, rs, f1, f2
	query = "SELECT DISTINCT CollectionName,SiteID FROM dbo.v_Collections ORDER BY CollectionName"
	Set conn = CreateObject("ADODB.Connection")
	'On Error Resume Next
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	If err.Number <> 0 Then
		CMWT_STOP err.Number & ": " & err.Description
	End If
	'On Error GoTo 0
	Set cmd  = CreateObject("ADODB.Command")
	Set rs   = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	rs.MoveFirst
	If Not(rs.BOF And rs.EOF) Then
		Do Until rs.EOF
			f1 = Trim(rs.Fields("CollectionName").value)
			f2 = Trim(rs.Fields("SiteID").value)
			Response.Write "<option value=""" & f2 & """>" & f1 & "</option>"
			rs.MoveNext
		Loop
		rs.Close
	End If
	Set cmd = Nothing
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
End Sub

CMWT_NewPage "", "", ""
PageBackLink = "cmtools.asp"
PageBackName = "CM Tools"
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<form name="form1" id="form1" method="post" action="">
<table class="tfx">
	<tr>
		<td class="v10 w300 vtop">
			<h4>Collections</h4>
			<select name="cn" id="cn" size="10" class="w300 v10" multiple="true">
				<% 
				CMWT_Collections
				%>
			</select>
		</td>
		<td class="v10 vtop">
			<h4>Action</h4>
			<input type="radio" name="r1" id="r1" value="A" /> Machine Policy Update<br/>
			<input type="radio" name="r1" id="r1" value="B" /> User Policy Update<br/>
			<input type="radio" name="r1" id="r1" value="C" /> Group Policy Update<br/>
			<input type="radio" name="r1" id="r1" value="D" /> Restart<br/>
			<input type="radio" name="r1" id="r1" value="E" /> Shut Down<br/>
		</td>
		<td class="v10 vtop">
			<h4>Comment</h4>
			<textarea name="comm" id="comm" rows="10" cols="40"></textarea>
		</td>
		<td class="v10 vtop right">
			<h4></h4>
			<br/>
			<p><input type="reset" name="b0" id="b0" value="Clear Form" class="btx w140 h32" /></p>
			<p><input type="button" name="b2" id="b2" value="Cancel" class="btx w140 h32" onClick="document.location.href='tools.asp'" /></p>
			<p><input type="submit" name="b1" id="b1" value="Execute" class="btx w140 h32" /></p>
		</td>
	</tr>
</table>
		
<% CMWT_Footer() %>

</body>
</html>