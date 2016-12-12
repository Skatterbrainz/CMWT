<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... collections.asp
' lastupdate..... 12/07/2016
' description.... collections listing page
'-----------------------------------------------------------------------------
time1 = Timer
SortBy  = CMWT_GET("s", "Name")
KeySet  = CMWT_GET("ks", "1")
KeySet2 = CMWT_GET("k2", "0")
QueryOn = CMWT_GET("qq", "")
ObjPFX  = CMWT_GET("ch", "ALL")

If (KeySet = "1") Then
	PageTitle = "User Collections"
Else
	PageTitle = "Device Collections"
End If
qx = ""
Select Case KeySet2
	Case "1":
		PageTitle = PageTitle & " (Query Based)"
		qx = " AND dbo.v_Collection.CollectionID IN (SELECT DISTINCT CollectionID FROM dbo.v_CollectionRuleQuery) "
	Case "2":
		PageTitle = PageTitle & " (Direct Membership)"
		qx = " AND dbo.v_Collection.CollectionID IN (SELECT DISTINCT CollectionID FROM dbo.v_CollectionRuleDirect) "
End Select
PageBackLink = "assets.asp"
PageBackName = "Assets"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
menulist = "0=All,1=Query Based,2=Direct Membership"

Response.Write "<table class=""t2""><tr>"
For each m in Split(menulist,",")
	mset = Split(m,"=")
	mlink = "collections.asp?ks=" & KeySet & "&k2=" & mset(0)
	If KeySet2 = mset(0) Then
		Response.Write "<td class=""m22"">" & mset(1) & "</td>"
	Else
		Response.Write "<td class=""m11"" onClick=""document.location.href='" & mlink & "'"">" & mset(1) & "</td>"
	End If
Next
Response.Write "</tr></table>"

CMWT_CLICKBAR objPfx, "collections.asp?ks=" & KeySet & "&k2=" & KeySet2 & "&ch=" 

If ucase(objPFX) <> "ALL" Then

	query = "SELECT DISTINCT " & _
		"dbo.v_Collection.Name, " & _
		"dbo.v_Collection.CollectionID, " & _
		"dbo.v_Collection.Comment, " & _
		"dbo.v_Collection.MemberCount, " & _
		"COUNT(dbo.v_CollectionVariable.Name) AS Variables " & _
		"FROM dbo.v_Collection LEFT OUTER JOIN " & _
		"dbo.v_CollectionRuleQuery ON dbo.v_Collection.CollectionID = dbo.v_CollectionRuleQuery.CollectionID " & _
		"LEFT OUTER JOIN " & _
		"dbo.v_CollectionVariable ON dbo.v_Collection.CollectionID = dbo.v_CollectionVariable.CollectionID " & _
		"WHERE (dbo.v_Collection.CollectionType = " & KeySet & ") " & _
		"AND (dbo.v_Collection.Name LIKE '" & ObjPFX & "%') " & _
		qx & "GROUP BY " & _
		"dbo.v_Collection.CollectionID, " & _
		"dbo.v_Collection.Name, " & _
		"dbo.v_Collection.Comment,  " & _
		"dbo.v_Collection.MemberCount " & _
		"ORDER BY " & SortBy

Else

	query = "SELECT DISTINCT " & _
		"dbo.v_Collection.Name, " & _
		"dbo.v_Collection.CollectionID, " & _
		"dbo.v_Collection.Comment, " & _
		"dbo.v_Collection.MemberCount, " & _
		"COUNT(dbo.v_CollectionVariable.Name) AS Variables " & _
		"FROM dbo.v_Collection LEFT OUTER JOIN " & _
		"dbo.v_CollectionRuleQuery ON dbo.v_Collection.CollectionID = dbo.v_CollectionRuleQuery.CollectionID " & _
		"LEFT OUTER JOIN " & _
		"dbo.v_CollectionVariable ON dbo.v_Collection.CollectionID = dbo.v_CollectionVariable.CollectionID " & _
		"WHERE (dbo.v_Collection.CollectionType = " & KeySet & ") " & _
		qx & "GROUP BY " & _
		"dbo.v_Collection.CollectionID, " & _
		"dbo.v_Collection.Name, " & _
		"dbo.v_Collection.Comment,  " & _
		"dbo.v_Collection.MemberCount " & _
		"ORDER BY " & SortBy

End If 
Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "collections.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()

Response.Write "</body></html>"
%>
