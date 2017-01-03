<!-- #include file=_settings.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... _core.asp
' lastupdate..... 01/02/2017
' description.... CMWT core functions library
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
' sub-name: CMWT_NewPage
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_NewPage (OnLoadRef, MetaRedirect, RedirectDelay)
	Dim mr
	If MetaRedirect <> "" Then 
		If RedirectDelay <> "" Then
			mr = "<meta http-equiv=""refresh"" content=""" & RedirectDelay & ";url=" & MetaRedirect & """ />"
		Else
			mr = "<meta http-equiv=""refresh"" content=""1;url=" & MetaRedirect & """ />"
		End If
	Else
		mr = ""
	End If
	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN " & _
		" http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & _
		"<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">" & _
		"<head><meta charset=""utf-8"">" & _
		"<meta http-equiv=""Content-Language"" content=""en-us"" />" & _
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"" />" & _
		"<meta http-equiv=""Cache-Control"" content=""cache"""" />" & _
		"<meta name=""distribution"" content=""Global"" />" & _
		"<meta name=""revisit-after"" content=""1 days"" />" & _
		"<meta name=""robots"" content=""follow, index, noodp, noydir"" />" & _
		"<meta name=""description"" content="""" />" & _
		"<meta name=""abstract"" content="""" />" & _
		"<meta name=""author"" content=""David M. Stein"" />" & _
		"<meta name=""copyright"" content=""David M. Stein"" />" & _
		"<meta name=""keywords"" content="""" />" & _
		"<title>CMWT: " & PageTitle & "</title>" & _
		"<link rel=""stylesheet"" type=""text/css"" href=""default.css"" />" & _
		"<link rel=""shortcut icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
		"<link rel=""icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
		"<script src=""_cmwt.js""></script>" & mr & _
		"</head>"
	If OnLoadRef <> "" Then 
		Response.Write "<body onLoad=""" & OnLoadRef & """>"
	Else
		If mr <> "" Then 
			Response.Write "<body><table class=""tfx"">" & _
				"<tr><td class=""v10 ctr pad66 h300"">" & _
				"<h2>Please wait: " & Caption & "...</h2>" & _
				"<img src=""images/nipple.GIF"" border=""0"" alt="""" />" & _
				"</td></tr></table></body></html>"
		Else
			Response.Write "<body>"
		End If
	End If
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_PageRedirect
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_PageRedirect (Target, Delay)
	Dim mr
	If Target = "" Then 
		CMWT_STOP "CMWT_PageRedirect: no target URL was specified."
	End If
	If Target <> "" Then 
		If Delay <> "" Then
			mr = "<meta http-equiv=""refresh"" content=""" & Delay & ";url=" & Target & """ />"
		Else
			mr = "<meta http-equiv=""refresh"" content=""1;url=" & Target & """ />"
		End If
	Else
		mr = ""
	End If
	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN " & _
		" http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & _
		"<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">" & _
		"<head><meta charset=""utf-8"" />" & _
		"<meta http-equiv=""Content-Language"" content=""en-us"" />" & _
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"" />" & _
		"<meta http-equiv=""Cache-Control"" content=""cache"""" />" & _
		"<meta name=""revisit-after"" content=""1 days"" />" & _
		"<meta name=""robots"" content=""follow, index, noodp, noydir"" />" & _
		"<title>CMWT: " & PageTitle & "</title>" & _
		"<link rel=""stylesheet"" type=""text/css"" href=""ripple.css"" />" & _
		"<link rel=""shortcut icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
		"<link rel=""icon"" href=""./favicon.ico"" type=""image/x-icon"">" & _
		"<script src=""_cmwt.js""></script>" & _
		"<meta http-equiv=""refresh"" content=""" & Delay & ";url=" & Target & """ />" & _
		"</head>" & _
		"<body>" & _
		"<table class=""tfx"">" & _
		"<tr><td class=""v10 ctr pad66 h300"">" & _
		"<h2>Please wait: " & Caption & "...</h2>" & _
		"<div class='uil-ripple-css' style='transform:scale(0.4);'><div></div><div></div></div>" & _
		"</td></tr></table></body></html>"
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_FOOTER
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_FOOTER ()
	Dim ltime
	ltime = CMWT_LOADTIME()
	If ltime <> "" Then
		ltime = " - Page load: " & ltime & " seconds"
	End If
	Response.Write "<br/><table class=""tfx""><tr><td class=""footer"">" & _
		"Version: " & Application("CMWT_VERSION") & _
		" - Build: " & Application("CMWT_BUILD") & _
		" - Copyright &copy; " & DatePart("yyyy",Now) & ltime & _
		" - <a href=""about.asp"" title=""About"">About</a>" & _
		"</td></tr></table>"
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_PageFile
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_PageFile ()
	CMWT_PageFile = Replace(Lcase(Request.ServerVariables("URL")),"/cmwt/","")
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_PageHeading
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_PageHeading (caption, MoreButtons)
	Response.Write "<table class=""tfx"">" & _
		"<tr><td><h2>" & caption & "</h2></td>" & _
		"<td class=""v10 w500 right"">"
	If MoreButtons <> "" Then 
		Response.Write MoreButtons
	End If
	Response.Write "<input type=""button"" id=""pp"" name=""pp"" value=""Print"" " & _
		"class=""btx w150 h32"" onClick=""javascript:print();"" title=""Print"" />" & _
		"</td></tr></table>"
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_UserName
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_UserName ()
	CMWT_UserName = Session("CMWT_USERNAME")
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_ADMIN
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_ADMIN ()
	If Session("CMWT_ADMIN") = "TRUE" Then
		CMWT_ADMIN = True
	ElseIf InStr(Application("CMWT_ADMINS"), CMWT_USERNAME()) > 0 OR Application("CMWT_ADMINS") = CMWT_USERNAME() Then
		Session("CMWT_ADMIN") = "TRUE"
		CMWT_ADMIN = True
	End If
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_MENULIST
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_MENULIST (def, baseURL)
	Dim mx, mz, m1, m2, lnk
	Response.Write "<select name=""mbar"" id=""mbar"" size=""1"" class=""w300 pad6"" " & _
		"onChange=""if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }"">"
	For each mx in Split(Application("CMWT_DMENULIST"),",")
		mz = Split(mx,":")
		m1 = mz(0) ' key name
		m2 = mz(1) ' description
		If CMWT_NotNullString(def) Then
			If Ucase(def) = Ucase(m1) Then
				Response.Write "<option value="""" selected>" & m1 & "</option>"
			Else
				lnk = baseURL & "&set=" & m1
				Response.Write "<option value=""" & lnk & """>" & m1 & "</option>"
			End If
		Else
			Response.Write "<option value=""" & m1 & """>" & m2 & "</option>"
		End If
	Next
	Response.Write "</select>"
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_MENUGROUP
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_MENUGROUP (code)
	Dim mx, mz, result : result = ""
	For each mx in Split(Application("CMWT_DMENULIST"), ",")
		mz = Split(mx,":")
		If Ucase(code) = Ucase(mz(0)) Then
			result = mz(1)
		End If
	Next
	CMWT_MENUGROUP = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_IsNullString
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_IsNullString (StringVal)
	If IsNull(StringVal) OR Trim(StringVal) = "" Then
		CMWT_IsNullString = True
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_NotNullString
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_NotNullString (StringVal)
	If Not CMWT_IsNullString(StringVal) Then
		CMWT_NotNullString = True
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_GET
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_GET (KeyName, DefaultValue)
	Dim result : result = ""
	result = Trim(Request.Form(KeyName))
	If result = "" Then
		result = Trim(Request.QueryString(KeyName))
	End If
	If result = "" Then
		result = DefaultValue
	End If
	CMWT_GET = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_GETX
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_GETX (x, v, m)
	Dim result : result = ""
	result = Trim(Request.Form(x))
	If result = "" Then
		result = Trim(Request.QueryString(x))
	End If
	If result = "" Then
		If v <> "" Then
			result = v
		End If
	End If
	If result = "" And m <> "" Then
		CMWT_STOP m
	End If
	CMWT_GETX = result
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_VALIDATE
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_VALIDATE (CheckValue, Message)
	If CheckValue = "" Then
		CMWT_STOP Message
	End If
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_STOP
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_STOP (Message)
	Response.Redirect "error.asp?m=" & Message
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_ButtonBar
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_ButtonBar (DataString, DefaultIndex, BaseWebLink)
	Dim bset, aset, mCode, mName
	Response.Write "<br/><div class=""tfx""><table class=""t2""><tr>"
	For each bset in Split(DataString, ",")
		aset = Split(bset,"=")
		mCode = aset(0)
		mName = aset(1)
		If mCode = DefaultIndex Then
			Response.Write "<td class=""td6 w140 ctr bgGray"">" & mName & "</td>"
		Else
			Response.Write "<td class=""td6 w140 ctr"" " & _
			"onMouseOver=""this.className='td6 w140 ctr ptr bgLightBlue'"" " & _
			"onMouseOut=""this.className='td6 w140 ctr'"" " & _
			"onClick=""document.location.href='" & BaseWebLink & mCode & "'"">" & mName & "</td>"
		End If
	Next
	Response.Write "</tr></table></div>"
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_CLICKBAR
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_CLICKBAR (DefaultChar, WebLink)
	Dim i
	Response.Write "<table class=""tfx""><tr>"
	If Ucase(DefaultChar) = "ALL" Then
		Response.Write "<td class=""m00"">All</td>"
		For i = 0 to 9
			Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & i & "'"" title=""Names beginning with " & i & """>" & i & "</td>"
		Next
		For i = ASC("A") To ASC("Z")
			Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & CHR(i) & "'"" title=""Names beginning with " & CHR(i) & """>" & CHR(i) & "</td>"
		Next
	Else
		Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & "ALL" & "'"" title=""All records"">All</td>"
		For i = 0 to 9
			If IsNumeric(DefaultChar) Then
				If CDbl(i) = CDbl(DefaultChar) Then
					Response.Write "<td class=""m00"">" & i & "</td>"
				Else
					Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & i & "'"" title=""Names beginning with " & i & """>" & i & "</td>"
				End If
			Else
				Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & i & "'"" title=""Names beginning with " & i & """>" & i & "</td>"
			End If
		Next
		For i = ASC("A") To ASC("Z")
			If i = ASC(DefaultChar) Then
				Response.Write "<td class=""m00"" title=""Names beginning with " & CHR(i) & """>" & CHR(i) & "</td>"
			Else
				Response.Write "<td class=""m01"" onClick=""document.location.href='" & WebLink & CHR(i) & "'"" title=""Names beginning with " & CHR(i) & """>" & CHR(i) & "</td>"
			End If
		Next
	End If
	Response.Write "</tr></table>"
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_Banner
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_Banner()
	Response.Write "<table class=""tfx"">" & _
		"<tr><td style=""cursor:pointer"" onClick=""document.location.href='./'"" title=""" & CMWT_PageTitle() & """>" & _
		"<img src=""./images/cmwt_banner3a.png"" border=""0"" alt="""" /></td>" & _
		"<td class=""w250 right"">" & _
		"<form name=""formS"" id=""formS"" method=""post"" action=""search.asp"">" & _
		"<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"">" & _
		"<tr><td class=""v10 bgWhite"">" & _
		"<input type=""text"" name=""q"" id=""q"" size=""32"" class=""sx1"" value=""Search"" onClick=""this.className='sx2';this.value=''""/>" & _
		"</td>" & _
		"<td style=""width:50px"">" & _
		"<input type=""submit"" name=""z1"" id=""z1"" value="""" class=""searchbutton"" />" & _
		"</td></tr></table></form></td></tr>" & _
		"<tr><td colspan=""2"">"
	Response.Write FormatDateTime(Now, vbLongDate)
	If PageTitle <> "" Then
		Response.Write " : <span style=""color:darkblue"">" & PageTitle & "</span>"
	End If
	Response.Write "</td></tr></table>"
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_ButtonBar
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_ButtonBar (DataString, DefaultIndex, BaseWebLink)
	Dim bset, aset, mCode, mName
	Response.Write "<br/><div class=""tfx""><table class=""t2""><tr>"
	For each bset in Split(DataString, ",")
		aset = Split(bset,"=")
		mCode = aset(0)
		mName = aset(1)
		If mCode = DefaultIndex Then
			Response.Write "<td class=""td6 w140 ctr bgGray"">" & mName & "</td>"
		Else
			Response.Write "<td class=""td6 w140 ctr"" " & _
			"onMouseOver=""this.className='td6 w140 ctr ptr bgBlue'"" " & _
			"onMouseOut=""this.className='td6 w140 ctr'"" " & _
			"onClick=""document.location.href='" & BaseWebLink & mCode & "'"">" & mName & "</td>"
		End If
	Next
	Response.Write "</tr></table></div>"
End Sub

'----------------------------------------------------------------
' function-name: CMWT_ButtonLinks
' function-desc: 
'----------------------------------------------------------------

function CMWT_ButtonLinks (DataString, DefaultIndex, BaseWebLink)
	Dim bset, aset, mCode, mName, tmp, result : result = ""
	For each bset in Split(DataString, ",")
		aset  = Split(bset,"=")
		mCode = aset(0)
		mName = aset(1)
		If mCode = DefaultIndex Then
			tmp = "<input type=""button"" id=""bbx" & mCode & """ name=""bbx" & mCode & """ " & _
				"value=""" & mName & """ class=""m22"" />"
		Else
			tmp = "<input type=""button"" id=""bbx" & mCode & """ name=""bbx" & mCode & """ " & _
				"value=""" & mName & """ class=""m11"" " & _
				"onClick=""document.location.href='" & BaseWebLink & mCode & "'"" " & _
				"/>"
		End If
		result = result & tmp
	Next
	CMWT_ButtonLinks = result
End function

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_QUERY
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_QUERY (dsn, query)
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open dsn
	Set cmd = Server.CreateObject("ADODB.Command")
	Set rs  = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If err.Number <> 0 Then
		response.write "[cmwt_db_query] error: " & err.Number & "<br/>description: " & err.Description
		response.end
	End If
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_OPEN
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_OPEN (dsn)
	Set conn = Server.CreateObject("ADODB.Connection")
	On Error Resume Next
	conn.ConnectionTimeOut = 5
	conn.Open dsn
	If err.Number <> 0 Then
		CMWT_STOP err.Number & ": " & err.Description
	End If
	On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_CLOSE
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_CLOSE ()
	On Error Resume Next
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
End Sub

'----------------------------------------------------------------
' function-name: CMWT_DB_ROWCOUNT
' function-desc: 
'----------------------------------------------------------------

Function CMWT_DB_ROWCOUNT (query)
	Dim result : result = 0
	CMWT_RSQUERY query
	result = rs.Fields("QTY").value
	CMWT_RSCLEAR()
	CMWT_DB_ROWCOUNT = result
End Function

'----------------------------------------------------------------
' sub-name: CMWT_RSQUERY
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_RSQUERY (query)
	Set cmd = Server.CreateObject("ADODB.Command")
	Set rs  = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = conn
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If err.Number <> 0 Then
		response.write "[cmwt_rsquery] error: " & err.Number & "<br/>description: " & err.Description
		response.end
	End If
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_RSCLEAR
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_RSCLEAR ()
	On Error Resume Next
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_AutoLink
' function-desc: 
'-----------------------------------------------------------------------------

function CMWT_AutoLink (ColumnName, LinkVal)
	dim result : result = LinkVal
	if CMWT_NotNullString(LinkVal) then
		select case Ucase(ColumnName)
			Case "COMPUTER","COMPUTERNAME","NETBIOSNAME","DEVICENAME","VMHOST":
				result = "<a href=""device.asp?cn=" & LinkVal & """ title=""Details for " & LinkVal & """>" & LinkVal & "</a>"
			Case "ADSITENAME":
				result = "<a href=""cmadsite.asp?sn=" & LinkVal & """ title=""Computers in Site " & LinkVal & """>" & LinkVal & "</a>"
			Case "ASSIGNMENTID":
				result = "<a href=""deptype.asp?id=" & LinkVal & """ title=""Deployment Type Details"">" & LinkVal & "</a>"
			Case "ADSITE":
				result = "<a href=""adsite.asp?sn=" & LinkVal & """ title=""Computers in AD Site " & LinkVal & """>" & LinkVal & "</a>"
			Case "USERNAME","USERID","SAMACCOUNTNAME","LOGONNAME":
				result = "<a href=""aduser.asp?uid=" & LinkVal & """ title=""User Account: " & LinkVal & """>" & LinkVal & "</a>"
			Case "GROUPNAME","GROUPNAME2":
				result = "<a href=""adgroup.asp?gn=" & LinkVal & """ title=""Details for " & LinkVal & """>" & LinkVal & "</a>"
			Case "MODEL0","MODELNAME","MODEL":
				result = "<a href=""model.asp?m=" & LinkVal & """ title=""Computers with model: " & LinkVal & """>" & LinkVal & "</a>"
			Case "MANUFACTURER":
				result = "<a href=""mfr.asp?mfr=" & LinkVal & """ title=""Devices by " & LinkVal & """>" & LinkVal & "</a>"
			Case "WINDOWS","WINDOWSTYPE","WINDOWSNAME","OSNAME":
				result = "<a href=""os.asp?on=" & LinkVal & """ title=""Computers with: " & LinkVal & """>" & LinkVal & "</a>"
			Case "OSCAPTION":
				result = "<a href=""os.asp?on=" & LinkVal & """ title=""Show Devices with " & LinkVal & """>" & LinkVal & "</a></td>"
			Case "CLIENT","ISACTIVE","ISPXE":
				result = CMWT_YESNO(LinkVal, True)
			Case "MEMORY":
				result = CMWT_KB2GB(LinkVal) & " GB"
			Case "SIZE":
				result = LinkVal & " GB"
			Case "UID":
				result = "<a href=""update.asp?id=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "COMPLIANT":
				result = LinkVal & "%"
			Case "DISCOVERY":
				result = "<a href=""discovery.asp?dm=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "ROLENAME":
				result = "<a href=""cmrole.asp?rn=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "PUBLISHER0","PUBLISHER","VENDOR","VENDORNAME":
				result = "<a href=""vendorapps.asp?vn=" & LinkVal & """ title=""Products by " & LinkVal & """>" & LinkVal & "</a>"
			Case "DISPLAYNAME0","PRODUCTNAME":
				result = "<a href=""app.asp?pn=" & Server.URLEncode(LinkVal) & """ title=""Installations of " & LinkVal & """>" & LinkVal & "</a>"
			Case "COLLECTIONID","SITEID","LIMITTOCOLLECTIONID":
				result = "<a href=""collection.asp?id=" & LinkVal & """ title=""Collection"">" & LinkVal & "</a>"
			Case "BOUNDARYGROUP":
				result = "<a href=""bgroup.asp?gn=" & LinkVal & """ title=""Boundary " & LinkVal & """>" & LinkVal & "</a>"
			Case "CHASSISTYPE":
				result = "<a href=""chassistype.asp?ct=" & LinkVal & """ title=""Computers of type " & LinkVal & """>" & LinkVal & "</a>"
			Case "DPSERVER":
				result = "<a href=""dpapplist.asp?dp=" & LinkVal & """ title=""View Applications..."">" & LinkVal & "</a>"
			Case "DPGROUP":
				result = "<a href=""dpgroup.asp?gn=" & LinkVal & """ title=""View Details..."">" & LinkVal & "</a>"
			Case "PACKAGEID","PKGID":
				result = "<a href=""package.asp?id=" & LinkVal & """ title=""View Details..."">" & LinkVal & "</a>"
			Case "TSPKGID":
				result = "<a href=""tasksequence.asp?id=" & LinkVal & """ title=""View Details..."">" & LinkVal & "</a>"
			Case "APPID":
				result = "<a href=""package.asp?k2=8&id=" & LinkVal & """ title=""Application Details..."">" & LinkVal & "</a>"
			Case "PUBLISHED","DISCOVERYENABLED","ISDEPLOYED","ISSUPERSEDED":
				result = CMWT_YESNO(LinkVal, True)
			Case "ISENABLED","ISACTIVE","ISDELETED","ISOBSOLETE","ISEXPIRED","ISHIDDEN","EULAEXISTS":
				result = CMWT_YESNO(LinkVal,True)
			Case "TASKNAME":
				result = "<a href=""cmtask.asp?tn=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "QUERYID":
				result = "<a href=""cmquery.asp?id=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "REPID":
				result = "<a href=""sqlrun.asp?id=" & LinkVal & """ title=""Run Report"">Run</a> . " & _
					"<a href=""sqlrepedit.asp?id=" & LinkVal & """ title=""Modify Report"">Edit</a> . " & _
					"<a href=""sqlrepdel.asp?id=" & LinkVal & """ title=""Delete Report"">Del</a>"
			Case "REPORTID":
				result = LinkVal & " ... " & _
					CMWT_IMG_LINK (True, "icon_add2", "icon_add1", "icon_add2", "reportrun.asp?id=" & LinkVal & "&rm=0", "Run Report") & " " & _
					CMWT_IMG_LINK (True, "icon_edit2", "icon_edit1", "icon_edit2", "reportedit.asp?id=" & LinkVal, "Edit Report") & " " & _
					CMWT_IMG_LINK (True, "icon_del2", "icon_del1", "icon_del2", "reportdel.asp?id=" & LinkVal, "Delete Report")
			Case "FILENAME":
				result = "<a href=""dupefiles.asp?cn=" & cn & "&fn=" & LinkVal & """ title=""View Instances"">" & LinkVal & "</a>"
			Case "COMPONENTNAME":
				result = "<a href=""ss2.asp?id=" & LinkVal & """ title=""View Details"">" & LinkVal & "</a>"
			Case "INFOURL":
				If PageTitle = "Software Update" Then
					result = "<a href=""" & LinkVal & """ target=""_blank"" title=""Open Link"">" & LinkVal & "</a>"
				Else
					result = "<a href=""" & LinkVal & """ target=""_blank"" title=""Open Link"">Link</a>"
				End If
			Case "PACKAGETYPE":
				Select Case LinkVal
					Case 0: result = "0 = Package"
					Case 8: result = "8 = Application"
					Case Else: result = LinkVal
				End Select
			Case "SHARETYPE":
				Select Case LinkVal
					Case 1: result = "1 = Common"
					Case 2: result = "2 = Specific"
					Case Else: result = LinkVal
				End Select
			Case "FORCEDDISCONNECTENABLED","IGNOREADDRESSSCHEDULE","ISREFERENCECOLLECTION","ISBUILTIN":
				result = CMWT_YESNO(LinkVal, True)
			Case "ACTIONINPROGRESS":
				Select Case LinkVal
					Case 0: result = "0 = None"
					Case 1: result = "1 = Update"
					Case 2: result = "2 = Add"
					Case 3: result = "3 = Delete"
					Case Else: result = LinkVal
				End Select
			Case "SEVERITYNAME","BULLETINID","ARTICLEID","EXPIRED","SUPERSEDED","DEPLOYED":
				result = "<a href=""updates.asp?fn=" & ColumnName & "&fv=" & LinkVal & """ title=""Filter on " & LinkVal & """>" & LinkVal & "</a>"
			Case "COMPONENT":
				result = "<a href=""compstatus.asp?fn=" & ColumnName & "&fv=" & LinkVal & """ title=""Filter on " & LinkVal & """>" & LinkVal & "</a>"
			Case "ADID":
				result = "<a href=""adr.asp?id=" & LinkVal & """ title=""Show Details"">" & LinkVal & "</a>"
		end select
	end if
	CMWT_AutoLink = result
end function

'-----------------------------------------------------------------------------
' function-name: CMWT_DB_ColumnJustify
' function-desc: 
'-----------------------------------------------------------------------------

function CMWT_DB_ColumnJustify (ColumnName)
	select case Ucase(ColumnName)
		case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
			CMWT_DB_ColumnJustify = "right"
		case "STATUS","TRUSTS","DOMAINS","ADSITES","SUBNETS","LASTDISCOVERYTIME":
			CMWT_DB_ColumnJustify = "right"
		case else:
			CMWT_DB_ColumnJustify = ""
	end select 
end function

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_TableGrid
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_TableGrid (rs, Caption, SortLink, AutoLink)
	if not (rs.BOF and rs.EOF) then 
		xrows = rs.RecordCount 
		xcols = rs.Fields.Count
		if CMWT_NotNullString(Caption) then 
			response.write "<h2 class=""tfx"">" & Caption & "</h2>"
		end if
		response.write "<table class=""tfx""><tr>"
		for i = 0 to xcols -1
			fn = rs.fields(i).name
			Select Case Ucase(fn)
				Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
					Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case "REPORTID":
					Response.Write "<td class=""td6 v10 bgGray w100 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case Else:
					Response.Write "<td class=""td6 v10 bgGray"">"
			End Select
			If CMWT_NotNullString(SortLink) Then
				Response.Write CMWT_SORTLINK(SortLink, fn, SortBy) & "</td>"
			Else
				Response.Write fn & "</td>"
			End If
		next
		Response.Write "</tr>"
		If AutoLink <> "" Then 
			alx = Split(AutoLink, "=")
			afn = alx(0)
			afl = alx(1)
		Else
			afn = ""
		End If
		Do Until rs.EOF
			Response.Write "<tr class=""tr1"">"
			For i = 0 to xcols-1
				fn = rs.Fields(i).Name
				fv = rs.Fields(i).Value
				If Ucase(afn) = Ucase(fn) Then
					fv = "<a href=""" & afl & "=" & fv & """>" & fv & "</a>"
				Else
					fv = CMWT_AutoLink (fn, fv)
				End If
				response.write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
			next
			rs.MoveNext
		Loop
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
			xrows & " rows returned"
		If filtered = True Then
			Response.Write " (Filtered Results)"
		End If
		Response.Write "</td></tr></table>"
	else
		If CMWT_NotNullString(Caption) Then
			Response.Write "<h2 class=""tfx"">" & Caption & "</h2>"
		End If
		Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching rows found</td></tr></table>"
	end if 
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_TableGrid2
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_TableGrid2 (rs, Caption, SortLink, AutoLink, FormLink)
	Dim xrows, xcols, fn, fv, alx, afn, afl, flx, fpn, fcn, i
	if not (rs.BOF and rs.EOF) then 
		xrows = rs.RecordCount 
		xcols = rs.Fields.Count
		if CMWT_NotNullString(Caption) then 
			Response.Write "<h2 class=""tfx"">" & Caption & "</h2>"
		end if
		Response.Write "<table class=""tfx""><tr>" & _
			"<td class=""td6 v10 ctr w30 bgGray"">&nbsp;</td>"
		for i = 0 to xcols -1
			fn = rs.fields(i).name
			Select Case Ucase(fn)
				Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
					Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case Else:
					Response.Write "<td class=""td6 v10 bgGray"">"
			End Select
			If CMWT_NotNullString(SortLink) Then
				Response.Write CMWT_SORTLINK(SortLink, fn, SortBy) & "</td>"
			Else
				Response.Write fn & "</td>"
			End If
		next
		Response.Write "</tr>"
		If AutoLink <> "" Then 
			alx = Split(AutoLink, "=")
			afn = alx(0)
			afl = alx(1)
		Else
			afn = ""
		End If
		flx = Split(FormLink, "=")
		' form property name
		fpn = flx(0)
		' form recordset column name
		fcn = flx(1)
		Do Until rs.EOF
			Response.Write "<tr class=""tr1"">"
			Response.Write "<td class=""td6 v10 ctr"">" & _
				"<input type=""checkbox"" class=""CB1"" name=""" & fpn & """ id=""" & _
				fpn & """ value=""" & rs.Fields(fcn).value & """ /></td>"
			For i = 0 to xcols-1
				fn = rs.Fields(i).Name
				fv = rs.Fields(i).Value
				If Ucase(afn) = Ucase(fn) Then
					fv = "<a href=""" & afl & "=" & fv & """>" & fv & "</a>"
				Else
					fv = CMWT_AutoLink (fn, fv)
				End If
				Response.Write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
			next
			rs.MoveNext
		Loop
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols+1 & """>" & _
			xrows & " rows returned</td></tr></table>"
	else
		If CMWT_NotNullString(Caption) Then
			Response.Write "<h2 class=""tfx"">" & Caption & "</h2>"
		End If
		Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching rows found</td></tr></table>"
	end if 
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_TableGrid
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_TableGridFilter (rs, Caption, SortLink, AutoLink, ColumnSet, FilterLink)
	if not (rs.BOF and rs.EOF) then 
		xrows = rs.RecordCount 
		xcols = rs.Fields.Count
		if CMWT_NotNullString(Caption) then 
			response.write "<h2 class=""tfx"">" & Caption & "</h2>"
		end if
		response.write "<table class=""tfx""><tr>"
		for i = 0 to xcols -1
			fn = rs.fields(i).name
			Select Case Ucase(fn)
				Case "QTY","RECS","COUNT","MEMBERS","GROUPCOUNT","COMPUTERS","CLIENTS","COVERAGE":
					Response.Write "<td class=""td6 v10 bgGray w80 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case "REPORTID":
					Response.Write "<td class=""td6 v10 bgGray w100 " & CMWT_DB_ColumnJustify(fn) & """>"
				Case Else:
					Response.Write "<td class=""td6 v10 bgGray"">"
			End Select
			If CMWT_NotNullString(SortLink) Then
				Response.Write CMWT_SORTLINK(SortLink, fn, SortBy) & "</td>"
			Else
				Response.Write fn & "</td>"
			End If
		next
		if ColumnSet <> "" then 
			Response.Write "<tr>"
				For each csn in Split(colset,",")
					csx = Split(csn,"=")
					Response.Write "<td class=""pad6 v10 bgDarkGray"">"
					If csx(1) = 1 Then
						Response.Write "<input type=""text"" name=""ss" & csx(0) & """ id=""" & csx(0) & """ " & _
							"maxlength=""50"" class=""pad5 v10"" title=""Filter: " & csx(0) & """ />"
					End If
					Response.Write "</td>"
				Next
				Response.Write "</tr>"
		end if
		Response.Write "</tr>"
		Do Until rs.EOF
			Response.Write "<tr class=""tr1"">"
			For i = 0 to xcols-1
				fn = rs.Fields(i).Name
				fv = rs.Fields(i).Value
				fv = "<a href=""" & FilterLink & "?fn=" & fn & "&fv=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
				response.write "<td class=""td6 v10 " & CMWT_DB_ColumnJustify(fn) & """>" & fv & "</td>"
			next
			rs.MoveNext
		Loop
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray"" colspan=""" & xcols & """>" & _
			xrows & " rows returned"
		If Filtered = True Then
			Response.Write " (Filtered Results) - <a href=""" & FilterLink & """ title=""Show All"">Show All</a>"
		End If
		Response.Write "</td></tr></table>"
	else
		If CMWT_NotNullString(Caption) Then
			Response.Write "<h2 class=""tfx"">" & Caption & "</h2>"
		End If
		Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching rows found</td></tr></table>"
	end if 
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_TABLEROWGRID
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_TABLEROWGRID (objRS, CaptionText, SortLink, AutoLink)
	Dim fn, fv, xrows, xcols, i, arrX, alx, x1, x2
	If Not(objRS.BOF And objRS.EOF) Then
		xrows = objRS.RecordCount
		xcols = objRS.Fields.Count
		If CMWT_NotNullString(CaptionText) Then
			Response.Write "<h2 class=""tfx"">" & CaptionText & "</h2>"
		End If
		Response.Write "<table class=""tfx"">"
		Do Until objRS.EOF
			For i = 0 to xcols-1
				fn = objRS.Fields(i).Name
				fv = objRS.Fields(i).Value
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10 w180 bgGray"">" & fn & "</td>" & _
					"<td class=""td6 v10"">" & CMWT_AutoLink(fn, fv) & "</td></tr>"
			Next
		
			objRS.MoveNext
		Loop
		Response.Write "</table>"
	Else
		If CMWT_NotNullString(CaptionText) Then
			Response.Write "<h2 class=""tfx"">" & CaptionText & "</h2>"
		End If
		Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching rows found</td></tr></table>"
	End If
	
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_DB_TABLEROWGRIDFilter
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DB_TABLEROWGRIDFilter (objRS, CaptionText, FilterLink)
	Dim fn, fv, xrows, xcols, i, arrX, alx, x1, x2
	If Not(objRS.BOF And objRS.EOF) Then
		xrows = objRS.RecordCount
		xcols = objRS.Fields.Count
		If CMWT_NotNullString(CaptionText) Then
			Response.Write "<h2 class=""tfx"">" & CaptionText & "</h2>"
		End If
		Response.Write "<table class=""tfx"">"
		Do Until objRS.EOF
			For i = 0 to xcols-1
				fn = objRS.Fields(i).Name
				fv = objRS.Fields(i).Value
				fv = "<a href=""" & FilterLink & "?fn=" & fn & "&fv=" & fv & """ title=""Filter on " & fv & """>" & fv & "</a>"
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10 w180 bgGray"">" & fn & "</td>" & _
					"<td class=""td6 v10"">" & fv & "</td></tr>"
			Next
		
			objRS.MoveNext
		Loop
		Response.Write "</table>"
	Else
		If CMWT_NotNullString(CaptionText) Then
			Response.Write "<h2 class=""tfx"">" & CaptionText & "</h2>"
		End If
		Response.Write "<table class=""tfx""><tr class=""h100 tr1"">" & _
			"<td class=""td6 v10 ctr"">No matching rows found</td></tr></table>"
	End If
	
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_WMI_TABLEGRID
' sub-desc: 
'-----------------------------------------------------------------------------
' CMWT_WMI_TABLEGRID ".", "Name,DisplayName,StartMode,State", "Win32_Service", "Services", "DisplayName", "Name=service.asp?sn="

Sub CMWT_WMI_TABLEGRID (hostname, columns, className, caption, sortby, autolink)
	Dim cn, objWMIService, colItems, objItem, val, PropertyName, afx, afn, afl, rows, cols
	Response.Write "<h2 class=""tfx"">" & caption & "</h2>" & _
		"<table class=""tfx""><tr>"
	cols = Ubound(Split(columns,","))+1
	For each cn in Split(columns, ",")
		Response.Write "<td class=""td6 v10 bgGray"">" & _
			"<a href=""" & Request.ServerVariables("PATH_INFO") & _
			"?s=" & cn & """ title=""Sort Column"">" & cn & "</a></td>"
	Next
	Response.Write "</tr>"
	
	Set objWMIService = GetObject("winmgmts:\\" & hostname & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery("SELECT " & columns & " FROM " & className,,48)
	If CMWT_NotNullString(autolink) Then
		afx = Split(autolink,"=")
		afn = afx(0)
		afl = afx(1)
	Else
		afn = ""
		afl = ""
	End If
	rows = 0
	For Each objItem in colItems
		Response.Write "<tr class=""tr1"">"
		For each PropertyName in Split(columns, ",")
			val = objItem.Properties_.Item(PropertyName)
			If CMWT_NotNullString(afn) And Ucase(afn)=Ucase(PropertyName) Then
				val = "<a href=""" & afl & "=" & val & """>" & val & "</a>"
			End If
			Response.Write "<td class=""td6 v10"">" & val & "</td>"
		Next
		Response.Write "</tr>"
		rows = rows + 1
	Next
	Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""" & cols & """>" & _
		rows & " services found</td></tr></table>"
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_TABLE_GRAPH2
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_TABLE_GRAPH2 (subcount, tcount)
	Dim pct
	If tcount > 0 and subcount > 0 Then
		pct = subcount / tcount
	Else
		pct = 0
	End If
	If pct > .5 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgGreen v8 cBlack"" width=""" & FormatPercent(pct,1) & """>" & _
			subcount & " (" & FormatPercent(pct,0) & ")</td>" & _
			"<td class=""pad6 v8""> </td></tr></table>"
	ElseIf pct > .15 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgBlue v8 cBlack"" width=""" & FormatPercent(pct,1) & """>" & _
			subcount & " (" & FormatPercent(pct,0) & ")</td>" & _
			"<td class=""pad6 v8""> </td></tr></table>"
	ElseIf pct > .08 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgLightOrange v8 cBlack"" width=""18%""> </td>" & _
			"<td class=""pad6 v8"">" & subcount & " (" & FormatPercent(pct,0) & ")</td></tr></table>"
	Else
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgLightOrange v8"" width=""10%""> </td>" & _
			"<td class=""pad6 v8"">" & subcount & " (" & FormatPercent(pct,0) & ")</td></tr></table>"
	End If
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_YESNO
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_YESNO (intVal, ExplicitNo)
	If CMWT_IsNullString(intVal) THen
		CMWT_YESNO = "NO"
	ElseIf intVal = 1 Then
		CMWT_YESNO = "YES"
	Else
		If ExplicitNo = True Then
			CMWT_YESNO = "NO"
		Else
			CMWT_YESNO = ""
		End If
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_SORTLINK
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_SORTLINK (BaseURL, ColName, DefSort)
	If Ucase(Trim(DefSort)) = Ucase(Trim(ColName)) Then
		If InStr(BaseURL,"?") > 0 Then
			CMWT_SORTLINK = "<a href=""" & BaseURL & "&s=" & ColName & " desc"" title=""Sort by " & ColName & " (descending)"">" & _
				"<img src=""images/sortdn.png"" border=""0""> " & ColName & "</a>"
		Else
			CMWT_SORTLINK = "<a href=""" & BaseURL & "?s=" & ColName & " desc"" title=""Sort by " & ColName & " (descending)"">" & _
				"<img src=""images/sortdn.png"" border=""0""> " & ColName & "</a>"
		End If
	Else
		If InStr(BaseURL,"?") > 0 Then
			CMWT_SORTLINK = "<a href=""" & BaseURL & "&s=" & ColName & """ title=""Sort by " & ColName & """>" & _
				"<img src=""images/sortup.png"" border=""0""> " & ColName & "</a>"
		Else
			CMWT_SORTLINK = "<a href=""" & BaseURL & "?s=" & ColName & """ title=""Sort by " & ColName & """>" & _
				"<img src=""images/sortup.png"" border=""0""> " & ColName & "</a>"
		End If
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_Get_CM_Property
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_Get_CM_Property (TableName, KeyName, KeyValue, ReturnCol)
	Dim conn, cmd, rs, query, clsx, result : result = ""
	query = "SELECT TOP 1 " & ReturnCol & " AS X " & _
		"FROM dbo." & TableName & _
		" WHERE (" & KeyName & "='" & KeyValue & "')"
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	On Error Resume Next
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	If err.Number <> 0 Then
		CMWT_STOP err.Number & ": " & err.Description
	End If
	clsx = True
	On Error GoTo 0
	
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
		result = rs.Fields("X").value
	End If
	rs.Close
	conn.Close
	CMWT_Get_CM_Property = result
End Function 

'-----------------------------------------------------------------------------
' function-name: CMWT_AppendQuery
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_AppendQuery (Q, KName, KValue, KMatch)
	Dim result
	Select Case Lcase(KMatch)
		Case "gt": result = Q & " WHERE (" & KName & ">'" & KValue & "')"
		Case "le": result = Q & " WHERE (" & KName & "<='" & KValue & "')"
		Case "ge": result = Q & " WHERE (" & KName & ">='" & KValue & "')"
		Case "lt": result = Q & " WHERE (" & KName & "<'" & KValue & "')"
		Case "like": result = Q & " WHERE (" & KName & " LIKE '%" & KValue & "%')"
		Case Else: result = Q & " WHERE (" & KName & "='" & KValue & "')"
	End Select
	CMWT_AppendQuery = result
End Function 

'-----------------------------------------------------------------------------
' sub-name: CMWT_TABLE_GRAPH
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_TABLE_GRAPH (pct, tcount, subcount)
	If subcount > 0 Then
		Response.Write "<table class=""t1"">" & _
			"<tr><td class=""td4 bgLightOrange v8 ctr"">" & tcount & "</td>" & _
			"<td class=""td4 bgLightGreen v8 ctr"" style=""width:" & pct & """>" & _
			subcount & "</td></tr></table>"
	Else
		Response.Write "<table class=""t1"">" & _
			"<tr><td class=""td6 bgLightOrange v8 ctr"">" & subcount & "</td>" & _
			"</tr></table>"
	End If
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_TABLE_GRAPH2
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_TABLE_GRAPH2 (subcount, tcount)
	Dim pct
	If tcount > 0 and subcount > 0 Then
		pct = subcount / tcount
	Else
		pct = 0
	End If
	If pct > .5 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgGreen v8"" width=""" & FormatPercent(pct,1) & """>" & subcount & " (" & FormatPercent(pct,0) & ")</td>" & _
			"<td class=""pad6 v8""> </td></tr></table>"
	ElseIf pct > .15 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgBlue v8"" width=""" & FormatPercent(pct,1) & """>" & subcount & " (" & FormatPercent(pct,0) & ")</td>" & _
			"<td class=""pad6 v8""> </td></tr></table>"
	ElseIf pct > .08 Then
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgOrange v8"" width=""18%""> </td>" & _
			"<td class=""pad6 v8"">" & subcount & " (" & FormatPercent(pct,0) & ")</td></tr></table>"
	Else
		Response.Write "<table class=""t1x"">" & _
			"<tr><td class=""pad6 bgOrange v8"" width=""10%""> </td>" & _
			"<td class=""pad6 v8"">" & subcount & " (" & FormatPercent(pct,0) & ")</td></tr></table>"
	End If
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_LOADTIME
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_LOADTIME ()
	If CMWT_NotNullString(time1) Then
		CMWT_LOADTIME = Round(Timer - CDBL(time1),2)
	Else
		CMWT_LOADTIME = 0
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_CM_IsCollectionMember
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_IsCollectionMember (c, ResourceName, CollectionID)
	Dim query, cmd, rs, result
	query = "SELECT TOP 1 CollectionID, ResourceID, Name " & _
		"FROM dbo.v_FullCollectionMembership " & _
		"WHERE (CollectionID='" & CollectionID & "') AND " & _
		"(Name='" & ResourceName & "')"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = True
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_CM_IsCollectionMember = result
End Function

'-----------------------------------------------------------------------------
' function-name: DPMS_CM_AddCollectionMember
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_AddCollectionMember (CollectionID, MachineName)
	Dim objLocator, objSMS, instColl, colNewResources,strNewResourceID, insNewResource, instDirectRule
	Dim result : result = 0
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")     
	Set objSMS = objLocator.ConnectServer(strServer, "root/SMS/site_" + Application("CM_SITECODE")) 
	objSMS.Security_.ImpersonationLevel = 3 
	Set instColl = objSMS.Get("SMS_Collection.CollectionID='" & CollectionID & "'")
	If Instcoll.Name <> "" Then
		Set colNewResources = objSMS.ExecQuery("SELECT ResourceId FROM SMS_R_System WHERE NetbiosName ='" & MachineName & "'")  
		strNewResourceID = 0       
		For each insNewResource in colNewResources 
			strNewResourceID = insNewResource.ResourceID 
		Next
		If strNewResourceID <> 0 Then
			Set instDirectRule = objSMS.Get("SMS_CollectionRuleDirect").SpawnInstance_ () 
			instDirectRule.ResourceClassName = "SMS_R_System"  
			instDirectRule.ResourceID = strNewResourceID 
			instDirectRule.RuleName = MachineName  
			instColl.AddMembershipRule instDirectRule, SMSContext 
			instColl.RequestRefresh False 
			if err.number <> 0 then
				result = Abs(err.number)
			else
				result = 1
			end if
		End If 
	End If
	CMWT_CM_AddCollectionMember = result
End Function

Function CMWT_Add_CmCollectionMember (strCompName, strCollID)
	Dim objSWbemLocator, objSWbemServices, ProviderLoc, Location, query
	Dim colCompResourceID, strNewResourceID, insCompResource
	Dim instColl, instDirectRule
	On Error Resume Next
	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
	Set objSWbemServices= objSWbemLocator.ConnectServer(Application("CMWT_SiteServer"),"root\sms") 
	 
	Set ProviderLoc = objSWbemServices.InstancesOf("SMS_ProviderLocation") 
	For Each Location In ProviderLoc 
		If Location.ProviderForLocalSite = True Then 
			Set objSWbemServices = objSWbemLocator.ConnectServer _ 
				(Location.Machine, "root\sms\site_" + Location.SiteCode) 
		End If 
	Next 

	query = "SELECT ResourceID FROM SMS_R_System WHERE NetbiosName='" & strCompName & "'"

	Set colCompResourceID = objSWbemServices.ExecQuery(query) 
	 
	For each insCompResource in colCompResourceID 
		strNewResourceID = insCompResource.ResourceID 
	Next 
	 
	Set instColl = objSWbemServices.Get("SMS_Collection.CollectionID=""" & strCollID & """") 
	Set instDirectRule = objSWbemServices.Get("SMS_CollectionRuleDirect").SpawnInstance_() 
	 
	instDirectRule.ResourceClassName = "SMS_R_System" 
	instDirectRule.ResourceID = strNewResourceID 
	instDirectRule.RuleName = strComputerName 
	instColl.AddMembershipRule instDirectRule
	CMWT_Add_CmCollectionMember = err.Number
End Function

'-----------------------------------------------------------------------------
' function-name: DPMS_CM_RemoveCollectionMember
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_RemoveCollectionMember (CollectionID, MachineName)
	Dim objLocator, objSMS, instColl, colNewResources,strNewResourceID, insNewResource, instDirectRule
	Dim result : result = 0
	Set objLocator = CreateObject("WbemScripting.SWbemLocator")     
	Set objSMS = objLocator.ConnectServer(strServer, "root/SMS/site_" + Application("CM_SITECODE")) 
	objSMS.Security_.ImpersonationLevel = 3 
	Set instColl = objSMS.Get("SMS_Collection.CollectionID='" & CollectionID & "'") 
	If Instcoll.Name <> "" Then
		Set colNewResources = objSMS.ExecQuery("SELECT ResourceId FROM SMS_R_System WHERE NetbiosName ='" & MachineName & "'")  
		strNewResourceID = 0       
		For each insNewResource in colNewResources 
			strNewResourceID = insNewResource.ResourceID 
		Next
		If strNewResourceID <> 0 Then
			Set instDirectRule = objSMS.Get("SMS_CollectionRuleDirect").SpawnInstance_ () 
			instDirectRule.ResourceClassName = "SMS_R_System"  
			instDirectRule.ResourceID = strNewResourceID 
			instDirectRule.RuleName = MachineName  
			instColl.DeleteMembershipRule instDirectRule, SMSContext 
			instColl.RequestRefresh False 
			if err.number <> 0 then
				result = Abs(err.number)
			else
				result = 1
			end if
		End If 
	End If
	CMWT_CM_RemoveCollectionMember = result
End Function

Function CMWT_Remove_CmCollectionMember (strCompName, strCollID)
	Dim objSWbemLocator, objSWbemServices, ProviderLoc, Location, query
	Dim colCompResourceID, strNewResourceID, insCompResource
	Dim instColl, instDirectRule
	On Error Resume Next
	Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
	Set objSWbemServices= objSWbemLocator.ConnectServer(Application("CMWT_SiteServer"),"root\sms") 
	 
	Set ProviderLoc = objSWbemServices.InstancesOf("SMS_ProviderLocation") 
	For Each Location In ProviderLoc 
		If Location.ProviderForLocalSite = True Then 
			Set objSWbemServices = objSWbemLocator.ConnectServer _ 
				(Location.Machine, "root\sms\site_" + Location.SiteCode) 
		End If 
	Next 

	query = "SELECT ResourceID FROM SMS_R_System WHERE NetbiosName='" & strCompName & "'"

	Set colCompResourceID = objSWbemServices.ExecQuery(query) 
	 
	For each insCompResource in colCompResourceID 
		strNewResourceID = insCompResource.ResourceID 
	Next 
	 
	Set instColl = objSWbemServices.Get("SMS_Collection.CollectionID=""" & strCollID & """") 
	Set instDirectRule = objSWbemServices.Get("SMS_CollectionRuleDirect").SpawnInstance_() 
	 
	instDirectRule.ResourceClassName = "SMS_R_System" 
	instDirectRule.ResourceID = strNewResourceID 
	instDirectRule.RuleName = strComputerName 
	instColl.DeleteMembershipRule instDirectRule
	CMWT_Remove_CmCollectionMember = err.Number
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_CM_ListCollections
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_CM_ListCollections (c, default, colltype, filterlist)
	Dim query, cmd, rs, x1, x2
	If filterlist <> "" Then
		query = "SELECT CollectionID, Name, CollectionType " & _
			"FROM CM_" & Application("CM_SITECODE") & ".dbo.v_Collection " & _
			"WHERE (CollectionID NOT IN " & _
			"(SELECT DISTINCT CollectionID FROM CM_" & Application("CM_SITECODE") & ".dbo.v_CollectionRuleQuery)) " & _
			"AND CollectionType=" & colltype & _
			" AND CollectionID NOT IN ('" & Replace(filterlist,",","','") & "') " & _
			" ORDER BY Name"
	Else
		query = "SELECT CollectionID, Name, CollectionType " & _
			"FROM CM_" & Application("CM_SITECODE") & ".dbo.v_Collection " & _
			"WHERE (CollectionID NOT IN " & _
			"(SELECT DISTINCT CollectionID FROM CM_" & Application("CM_SITECODE") & ".dbo.v_CollectionRuleQuery)) " & _
			"AND CollectionType=" & colltype & _
			" ORDER BY Name"
	End If
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		Do Until rs.EOF
			x1 = rs.Fields("CollectionID").value
			x2 = rs.Fields("Name").value
			Response.Write "<option value=""" & x1 & """>" & x2 & "</option>"
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_CM_ResourceDirectCollections
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_ResourceDirectCollections (c, ResourceName)
	Dim query, cmd, rs, result, x2
	result = ""
	query = "SELECT CollectionID " & _
		"FROM dbo.v_ClientCollectionMembers " & _
		"WHERE (Name = '" & ResourceName & "') AND " & _
		"(CollectionID NOT IN (SELECT DISTINCT CollectionID " & _
		"FROM dbo.v_CollectionRuleQuery))"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		Do Until rs.EOF
			x1 = rs.Fields("CollectionID").value
			If result <> "" Then
				result = result & "," & x1
			Else
				result = x1
			End If
			rs.MoveNext
		Loop
	End If
	CMWT_CM_ResourceDirectCollections = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_CM_CollectionExists
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_CollectionExists (c, CollectionName, CollectionType)
	Dim query, cmd, rs, result
	query = "SELECT CollectionID " & _
		"FROM CM_" & Application("CM_SITECODE") & ".dbo.v_Collection " & _
		"WHERE (Name = '" & CollectionName & "') AND (CollectionType=" & CollectionType & ")"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = True
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_CM_CollectionExists = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_CM_CollectionRuleType
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_CollectionRuleType (c, CollectionID)
	Dim query, cmd, rs, result
	query = "SELECT TOP 1 RuleName " & _
		"FROM dbo.v_COLLECTIONRULEQUERY WHERE CollectionID = '" & CollectionID & "'"
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = "QUERY"
	Else
		result = "DIRECT"
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_CM_CollectionRuleType = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_PrettySQL
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_PrettySQL (sqlString)
	CMWT_PrettySQL = _
		Replace( _
			Replace( _
				Replace(Ucase(sqlString), ",", ", "), "FROM", "<br/>FROM"), _
				"WHERE", "<br/>WHERE")
			
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_SHOW_QUERY
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_SHOW_QUERY ()
	Dim rlink, qlist, flist, rmod
	rlink = Request.ServerVariables("PATH_INFO")
	qlist = Trim(Request.QueryString())
	flist = Trim(Request.Form())
	If CMWT_NotNullString(qlist) Then
		rlink = rlink & "?" & qlist
		rmod = 1
	End If
	If CMWT_NotNullString(flist) Then
		If CMWT_IsNullString(qlist) Then
			rlink = rlink & "?" & flist
		Else
			rlink = rlink & "&" & flist
		End If
		rmod = 1
	End If
	If Right(rlink,1) = "&" Then
		rlink = Left(rlink, Len(rlink)-1)
	End If
	If Right(rlink,1) = "?" Then
		rlink = Left(rlink, Len(rlink)-1)
	End If
	If QueryOn = "1" Then
		Response.Write "<br/><div class=""tfx""><h3>T-SQL Statement</h3><table class=""tfx""><tr>" & _
			"<td class=""td6a v8 cYellow"">" & CMWT_PrettySQL(query) & "</td></tr></table><br/>" & _
			"<input type=""button"" name=""bq"" id=""bq"" class=""w150 h32 btx"" " & _
			"value=""Hide Query"" onClick=""document.location.href='" & Replace(rlink,"qq=1","") & "'"" />" & _
			"</div>"
	Else
		If CMWT_Hide_QueryLink <> True Then
			If CMWT_ADMIN() Then
				If rmod = 1 Then
					rlink = rlink & "&qq=1"
				Else
					rlink = rlink & "?qq=1"
				End If
				Response.Write "<br/><div class=""tfx"">" & _
					"<input type=""button"" name=""bq"" id=""bq"" class=""w150 h32 btx"" " & _
					"value=""Show Query"" onClick=""document.location.href='" & rlink & "'"" />" & _
					"</div>"
			Else
				Response.Write "<br/><div class=""tfx"">" & _
					"<input type=""button"" name=""bq"" id=""bq"" class=""w150 h32 btx"" " & _
					"value=""Show Query"" disabled=""true"" />" & _
					"</div>"
			End If
		End If
	End If
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_IMG_LINK
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_IMG_LINK (IsEnabled, icon_off, icon_on, icon_disabled, url, tip)
	If IsEnabled = True Then
		CMWT_IMG_LINK = "<img src=""images/" & icon_off & ".png"" border=""0"" " & _
			"onMouseOver=""this.src='images/" & icon_on & ".png'"" " & _
			"onMouseOut=""this.src='images/" & icon_off & ".png'"" " & _
			"title=""" & tip & """ " & _
			"onClick=""document.location.href='" & url & "'"" style=""cursor:pointer"" />"
	Else
		CMWT_IMG_LINK = "<img src=""images/" & icon_disabled & ".png"" border=""0"" title=""" & tip & """ />"
	End If
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_IMGLINK
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_IMGLINK (icon_off, icon_on, url, tip)
	Response.Write "<img src=""images/" & icon_off & ".png"" border=""0"" " & _
		"onMouseOver=""this.src='images/" & icon_on & ".png'"" " & _
		"onMouseOut=""this.src='images/" & icon_off & ".png'"" " & _
		"title=""" & tip & """ " & _
		"onClick=""document.location.href='" & url & "'"" style=""cursor:pointer"" />"
End Sub

'-----------------------------------------------------------------------------
' sub-name: CMWT_IMGLINK2
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_IMGLINK2 (IsEnabled, icon_off, icon_on, icon_disabled, url, tip)
	If IsEnabled = True Then
		Response.Write "<img src=""images/" & icon_off & ".png"" border=""0"" " & _
			"onMouseOver=""this.src='images/" & icon_on & ".png'"" " & _
			"onMouseOut=""this.src='images/" & icon_off & ".png'"" " & _
			"title=""" & tip & """ " & _
			"onClick=""document.location.href='" & url & "'"" style=""cursor:pointer"" />"
	Else
		Response.Write "<img src=""images/" & icon_disabled & ".png"" border=""0"" title=""" & tip & """ />"
	End If
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_IsOnline
' function-desc: return TRUE if computer responds to a PING request
'-----------------------------------------------------------------------------
 
Function CMWT_IsOnline (strComputer)
	Dim objPing, objStatus, retval
	If strComputer <> "" Then
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
			ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address='" & strComputer & "'")
		For Each objStatus in objPing
			If Not(IsNull(objStatus.StatusCode)) And objStatus.StatusCode = 0 Then
				CMWT_IsOnline = True
			End If
		Next
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_KB2GB
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_KB2GB (intVal)
	If intVal > 0 Then
		CMWT_KB2GB = Round(intVal / 1024 / 1024, 2)
	Else
		CMWT_KB2GB = 0
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_MB2GB
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_MB2GB (intVal)
	If intVal > 0 Then
		CMWT_MB2GB = Round(intVal / 1024, 2)
	Else
		CMWT_MB2GB = 0
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_DN_ADSI
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_DN_ADSI (dnString)
	Dim tmp, result : result = ""
	For each tmp in Split(dnString, ",")
		If Ucase(Left(tmp,2)) = "OU" Then
			If result <> "" Then
				result = Mid(tmp,4) & "\" & result
			Else
				result = Mid(tmp,4)
			End If
		ElseIf Ucase(Left(tmp,2)) = "DC" Then
			If Lcase(Right(tmp,3)) <> "com" Then
				If result <> "" Then
					result = Ucase(Mid(tmp,4)) & "\" & result
				Else
					result = Ucase(Mid(tmp,4))
				End If
			End If
		End If
	Next
	CMWT_DN_ADSI = result
End Function

'-----------------------------------------------------------------------------
' sub-name: CMWT_DEVICE_TABLE
' sub-desc: 
'-----------------------------------------------------------------------------

Sub CMWT_DEVICE_TABLE (q)
	Dim found, xrows, xcols, fn, fv, i
	On Error Resume Next
	Response.Write "<table class=""tfx"">"
	CMWT_DB_QUERY Application("DSN_CMDB"), q
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
				fv = CMWT_AutoLink (fn, fv)
				Response.Write "<td class=""td6 v10"">" & fv & "</td>"
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
End Sub

'-----------------------------------------------------------------------------
' function-name: CMWT_HOST_BY_IP
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_HOST_BY_IP (c, IPAddress)
	Dim query, cmd, rs, result
	query = "SELECT TOP 1 DNSHostName0 " & _
		"FROM dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION " & _
		"WHERE IPAddress0 = '" & IPAddress & "'"
	Set cmd  = CreateObject("ADODB.Command")
	Set rs   = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = rs.Fields("DNSHostName0").value
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_HOST_BY_IP = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_IP_BY_HOSTNAME
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_IP_BY_HOSTNAME (c, HostName)
	Dim query, cmd, rs, result
	query = "SELECT TOP 1 IPAddress0 " & _
		"FROM dbo.v_GS_NETWORK_ADAPTER_CONFIGURATION " & _
		"WHERE DNSHostName0 = '" & HostName & "'"
	Set cmd  = CreateObject("ADODB.Command")
	Set rs   = CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	Set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	If Not(rs.BOF And rs.EOF) Then
		result = rs.Fields("IPAddress0").value
	End If
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	CMWT_IP_BY_HOSTNAME = result
End Function

'-----------------------------------------------------------------------------
' function-name: URLDecode
' function-desc: 
'-----------------------------------------------------------------------------

Function URLDecode (sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
    sOutput = REPLACE(sConvert, "+", " ")
    aSplit = Split(sOutput, "%")
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
    URLDecode = sOutput
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_BROWSER_TYPE
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_BROWSER_TYPE ()
	Dim result : result = ""
	Dim tmp : tmp = Request.ServerVariables("HTTP_USER_AGENT")
	If InStr(Ucase(tmp),"CHROME") > 0 Then
		result = "CHROME"
	ElseIf InStr(Ucase(tmp),"MSIE 7.0") > 0 Then
		result = "IE7"
	ElseIf InStr(Ucase(tmp),"TRIDENT") > 0 Then
		result = "IE"
	ElseIf InStr(Ucase(tmp),"SAFARI") > 0 Then
		result = "SAFARI"
	End If
	CMWT_BROWSER_TYPE = result
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_CN
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CN (FQDN)
	Dim tmp : tmp = ""
	If InStr(FQDN,",") > 0 Then
		tmp = Split(FQDN,",")
		CMWT_CN = tmp(0)
	Else
		CMWT_CN = FQDN
	End If
End Function

'-----------------------------------------------------------------------------
' function-name: CMWT_WMI_ChassisType
' function-desc: 
'-----------------------------------------------------------------------------

function CMWT_WMI_ChassisType (MatchValue)
	dim x, y, z, result
	x = "1=Virtual," & _
		 "2=Blade Server," & _
		 "3=Desktop," & _
		 "4=Low-Profile Desktop," & _
		 "5=Pizza-Box," & _
		 "6=Mini Tower," & _
		 "7=Tower," & _
		 "8=Portable," & _
		 "9=Laptop," & _
		 "10=Notebook," & _
		 "11=Hand-Held," & _
		 "12=Mobile Device in Docking Station," & _
		 "13=All-in-One," & _
		 "14=Sub-Notebook," & _
		 "15=Space Saving Chassis," & _
		 "16=Ultra Small Form Factor," & _
		 "17=Server Tower Chassis," & _
		 "18=Mobile Device in Docking Station," & _
		 "19=Sub-Chassis," & _
		 "20=Bus-Expansion chassis," & _
		 "21=Peripheral Chassis," & _
		 "22=Storage Chassis," & _
		 "23=Rack-Mounted Chassis," & _
		 "24=Sealed-Case PC"
	if IsNumeric(MatchValue) Then 
		for each y in Split(x,",")
			if split(y,"=")(0) = MatchValue then
				result = split(y,"=")(1)
				exit for
			end if
		next
	else
		for each y in Split(x,",")
			if split(y,"=")(1) = MatchValue then
				result = split(y,"=")(0)
				exit for
			end if
		next
	end if 
	CMWT_WMI_ChassisType = result
end function

'-----------------------------------------------------------------------------
' function-name: CMWT_CM_CHASSISTYPE
' function-desc: 
'-----------------------------------------------------------------------------

Function CMWT_CM_CHASSISTYPE (CTNum)
	Select Case CTNum
		Case  1:	CMWT_CM_CHASSISTYPE = "Virtual"
		Case  2:	CMWT_CM_CHASSISTYPE = "Blade Server"
		Case  3:	CMWT_CM_CHASSISTYPE = "Desktop"
		Case  4:	CMWT_CM_CHASSISTYPE = "Low-Profile Desktop"
		Case  5:	CMWT_CM_CHASSISTYPE = "Pizza-Box"
		Case  6:	CMWT_CM_CHASSISTYPE = "Mini Tower"
		Case  7:	CMWT_CM_CHASSISTYPE = "Tower"
		Case  8:	CMWT_CM_CHASSISTYPE = "Portable"
		Case  9:	CMWT_CM_CHASSISTYPE = "Laptop"
		Case 10:	CMWT_CM_CHASSISTYPE = "Notebook"
		Case 11:	CMWT_CM_CHASSISTYPE = "Hand-Held"
		Case 12:	CMWT_CM_CHASSISTYPE = "Mobile Device in Docking Station"
		Case 13:	CMWT_CM_CHASSISTYPE = "All-in-One"
		Case 14:	CMWT_CM_CHASSISTYPE = "Sub-Notebook"
		Case 15:	CMWT_CM_CHASSISTYPE = "Space Saving Chassis"
		Case 16:	CMWT_CM_CHASSISTYPE = "Ultra Small Form Factor"
		Case 17:	CMWT_CM_CHASSISTYPE = "Server Tower Chassis"
		Case 18:	CMWT_CM_CHASSISTYPE = "Mobile Device in Docking Station"
		Case 19:	CMWT_CM_CHASSISTYPE = "Sub-Chassis"
		Case 20:	CMWT_CM_CHASSISTYPE = "Bus-Expansion chassis"
		Case 21:	CMWT_CM_CHASSISTYPE = "Peripheral Chassis"
		Case 22:	CMWT_CM_CHASSISTYPE = "Storage Chassis"
		Case 23:	CMWT_CM_CHASSISTYPE = "Rack-Mounted Chassis"
		Case 24:	CMWT_CM_CHASSISTYPE = "Sealed-Case PC"
		Case Else: CMWT_CM_CHASSISTYPE = ""
	End Select
End Function

'----------------------------------------------------------------
' function-name: CMWT_CM_CHASSISNUM
' function-desc: 
'----------------------------------------------------------------

Function CMWT_CM_CHASSISNUM (strCT)
	Select Case strCT
		Case "Virtual": CMWT_CM_CHASSISNUM = 1
		Case "Blade Server": CMWT_CM_CHASSISNUM = 2
		Case "Desktop": CMWT_CM_CHASSISNUM = 3
		Case "Low-Profile Desktop": CMWT_CM_CHASSISNUM = 4
		Case "Pizza-Box": CMWT_CM_CHASSISNUM = 5
		Case "Mini Tower": CMWT_CM_CHASSISNUM = 6
		Case "Tower": CMWT_CM_CHASSISNUM = 7
		Case "Portable": CMWT_CM_CHASSISNUM = 8
		Case "Laptop": CMWT_CM_CHASSISNUM = 9
		Case "Notebook": CMWT_CM_CHASSISNUM = 10
		Case "Hand-Held": CMWT_CM_CHASSISNUM = 11
		Case "Mobile Device in Docking Station": CMWT_CM_CHASSISNUM = 12
		Case "All-in-One": CMWT_CM_CHASSISNUM = 13
		Case "Sub-Notebook": CMWT_CM_CHASSISNUM = 14
		Case "Space Saving Chassis": CMWT_CM_CHASSISNUM = 15
		Case "Ultra Small Form Factor": CMWT_CM_CHASSISNUM = 16
		Case "Server Tower Chassis": CMWT_CM_CHASSISNUM = 17
		Case "Mobile Device in Docking Station": CMWT_CM_CHASSISNUM = 18
		Case "Sub-Chassis": CMWT_CM_CHASSISNUM = 19
		Case "Bus-Expansion chassis": CMWT_CM_CHASSISNUM = 20
		Case "Peripheral Chassis": CMWT_CM_CHASSISNUM = 21
		Case "Storage Chassis": CMWT_CM_CHASSISNUM = 22
		Case "Rack-Mounted Chassis": CMWT_CM_CHASSISNUM = 23
		Case "Sealed-Case PC": CMWT_CM_CHASSISNUM = 24
		Case Else: CMWT_CM_CHASSISNUM = 0
	End Select
End Function

'----------------------------------------------------------------
' function-name: CMWT_CM_CollectionType
' function-desc: 
'----------------------------------------------------------------

Function CMWT_CM_CollectionType (intval)
	Select Case intVal
		Case 2: CMWT_CM_CollectionType = "DEVICE"
		Case 1: CMWT_CM_CollectionType = "USER"
		Case Else: CMWT_CM_CollectionType = intval
	End Select
End Function

'----------------------------------------------------------------
' function-name: CMWT_CM_RESOURCEID
' function-desc: 
'----------------------------------------------------------------

Function CMWT_CM_RESOURCEID (strName)
	Dim query, conn, cmd, rs, result : result = ""
	query = "SELECT TOP 1 ResourceID FROM dbo.v_R_System WHERE Name0='" & strName & "'"
	Set conn = Server.CreateObject("ADODB.Connection")
	Set cmd  = Server.CreateObject("ADODB.Command")
	Set rs   = Server.CreateObject("ADODB.Recordset")
	On Error Resume Next
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
	If err.Number = 0 Then
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		Set cmd.ActiveConnection = conn
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd
		If Not(rs.BOF And rs.EOF) Then
			result = rs.Fields("ResourceID").value
		End If
	End If
	conn.Close
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CM_RESOURCEID = result
End Function

'----------------------------------------------------------------
' function-name: LargeIntegerToDate
' function-desc: takes Microsoft LargeInteger value (Integer8) and returns according the date and time
'----------------------------------------------------------------

Function LargeIntegerToDate (value)
	Dim sho, timeshiftvalue, timeshift, i8high, i8low
    timeShift = 240
    'get the large integer into two long values (high part and low part)
    i8High = value.HighPart
    i8Low = value.LowPart
    If (i8Low < 0) Then
           i8High = i8High + 1 
    End If
    'calculate the date and time: 100-nanosecond-steps since 12:00 AM, 1/1/1601
    If (i8High = 0) And (i8Low = 0) Then 
        LargeIntegerToDate = #1/1/1601#
    Else 
        LargeIntegerToDate = #1/1/1601# + (((i8High * 2^32) + i8Low)/600000000 - timeShift)/1440 
    End If
End Function

'----------------------------------------------------------------
' sub-name: CMWT_DEBUG
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_DEBUG (stringval)
	If CMWT_GET("debug", "") = "1" Then
		Response.Write stringval
		Response.End
	End If
End Sub

'----------------------------------------------------------------
' function-name: Get_BitValues
' function-desc: 
'----------------------------------------------------------------

function Get_BitValues (intVal)
	Dim x, y, i, n, result : result = ""
	y = Array( 1, 2, 4, 8, 16, 32, 64 )
	x = Array( "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" )
	for i = 0 to UBound(x)
		n = 2^i
		if intVal and n Then
			if result <> "" then
				result = result & "," & x(i)
			else
				result = x(i)
			end if
		end if
	next
	Get_BitValues = result
end function

'----------------------------------------------------------------
' function-name: CMWT_CM_BuildName
' function-desc: 
'----------------------------------------------------------------

function CMWT_CM_BuildName (versionnumber)
	Select Case versionnumber
		Case "5.00.8458.1000": CMWT_CM_BuildName = "1610"
		Case "5.00.8412.1000": CMWT_CM_BuildName = "1606"
		Case "5.00.8355.1000": CMWT_CM_BuildName = "1602"
		Case "5.00.8325.1000": CMWT_CM_BuildName = "1511"
		Case "5.00.8239.1403": CMWT_CM_BuildName = "2012 R2 SP1 CU3"
		Case "5.00.8239.1406": CMWT_CM_BuildName = "2012 R2 SP1 CU4"
		Case "5.00.8239.1407": CMWT_CM_BuildName = "2012 R2 SP1 CU2+"
		Case "5.00.8239.1501": CMWT_CM_BuildName = "2012 R2 SP1 CU2+"
		Case "5.00.8239.1301": CMWT_CM_BuildName = "2012 R2 SP1 CU2"
		Case "5.00.8239.1203": CMWT_CM_BuildName = "2012 R2 SP1 CU1"
		Case "5.00.8239.1000": CMWT_CM_BuildName = "2012 R2 SP1"
		Case "5.00.7958.1203": CMWT_CM_BuildName = "2012 R2 CU1"
		Case "5.00.7958.1303": CMWT_CM_BuildName = "2012 R2 CU2"
		Case "5.00.7958.1401": CMWT_CM_BuildName = "2012 R2 CU3"
		Case "5.00.7958.1501": CMWT_CM_BuildName = "2012 R2 CU4"
		Case "5.00.7958.1604": CMWT_CM_BuildName = "2012 R2 CU5"
		Case "5.00.7958.1000": CMWT_CM_BuildName = "2012 R2 RTM"
		Case Else: CMWT_CM_BuildName = ""
	End Select 
end function

'----------------------------------------------------------------
' function-name: CMWT_CM_ObjectName
' function-desc: 
'----------------------------------------------------------------

function CMWT_CM_ObjectName (c, ObjectType, ObjectID)
	dim query, cmd, rs, result : result = ""
	query = ""
	select case ucase(ObjectType)
		case "COLLECTION":
			query = "SELECT TOP 1 Name AS X FROM dbo.v_Collection WHERE CollectionID='" & ObjectID & "'"
		case "PACKAGE","APPLICATION":
			query = "SELECT TOP 1 Name AS X FROM dbo.v_Packages WHERE PackageID='" & ObjectID & "'"
		case "DEVICE","RESOURCE":
			query = "SELECT TOP 1 Name0 AS X FROM dbo.v_R_System WHERE ResourceiD='" & ObjectID & "'"
	end select
	if query <> "" then
		on error resume next
		set cmd  = Server.CreateObject("ADODB.Command")
		set rs   = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = adUseClient
		rs.CursorType = adOpenStatic
		rs.LockType = adLockReadOnly
		set cmd.ActiveConnection = c
		cmd.CommandType = adCmdText
		cmd.CommandText = query
		rs.Open cmd
		if not(rs.BOF and rs.EOF) then
			result = rs.Fields("X").value
		end if
		rs.Close
		set rs = Nothing
		set cmd = Nothing
	end if 
	CMWT_CM_ObjectName = result
end function

'----------------------------------------------------------------
' function-name: CMWT_CM_ObjectName
' function-desc: 
'----------------------------------------------------------------

function CMWT_CM_ObjectProperty (c, TableName, KeyField, KeyVal, ReturnField)
	dim query, cmd, rs, result : result = ""
	query = "SELECT TOP 1 " & ReturnField & " FROM dbo." & TableName & " WHERE (" & KeyField & "='" & KeyVal & "')"
	on error resume next
	set cmd  = Server.CreateObject("ADODB.Command")
	set rs   = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseClient
	rs.CursorType = adOpenStatic
	rs.LockType = adLockReadOnly
	set cmd.ActiveConnection = c
	cmd.CommandType = adCmdText
	cmd.CommandText = query
	rs.Open cmd
	if not(rs.BOF and rs.EOF) then
		result = rs.Fields( ReturnField ).value
	end if
	rs.Close
	set rs = Nothing
	set cmd = Nothing
	CMWT_CM_ObjectProperty = result
end function

'----------------------------------------------------------------
' function-name: Get_WMI_Properties
' function-desc: 
' example: xdata = Get_WMI_Properties (".", "Win32_ComputerSystem", "Manufacturer,Model,SystemType")
'----------------------------------------------------------------

function Get_WMI_Properties (Computer, ClassName, PropertyNames)
	dim objWMI, colItems, objItem, PropertyName, row, result : result = ""
	on error resume next
	set objWMI = getobject("winmgmts:\\" & Computer & "\root\CIMV2")
	if err.Number = 0 then
		set colItems = objWMI.ExecQuery ( "SELECT * FROM " & ClassName,,48 )
		for each objItem in colItems
			for each PropertyName in split(PropertyNames, ",")
				val = objItem.Properties_.Item(PropertyName)
				row = PropertyName & "=" & val
				if result <> "" then
					result = result & vbCRLF & row
				else
					result = row
				end if
				row = ""
			next
		next
	else
		result = "ERROR: Access Denied!"
	end if
	Get_WMI_Properties = result
end function

'----------------------------------------------------------------
' function-name: CMWT_PageLink
' function-desc: 
'----------------------------------------------------------------

function CMWT_PageLink (CMCategory, CMID)
	Select Case Ucase(CMCategory)
		Case "COMPUTER":
			CMWT_PageLink = CMWT_GET("t","device.asp?cn=" & CMID & "&set=Notes")
		Case "COLLECTION":
			CMWT_PageLink = CMWT_GET("t","collection.asp?id=" & CMID & "&ks=5")
		Case "PACKAGE":
			CMWT_PageLink = CMWT_GET("t", "package.asp?id=" & CMID & "&ks=5")
		Case Else:
			CMWT_PageLink = "./"
	End Select
end function

'----------------------------------------------------------------
' sub-name: CMWT_LogEvent
' sub-desc: 
'----------------------------------------------------------------

sub CMWT_LogEvent (c, EventType, EventCategory, Comment)
	dim query, conn
	if Application("CMWT_ENABLE_LOGGING") = "TRUE" Then
		query = "INSERT INTO dbo.EventLog " & _
			"(EventType, EventCategory, EventOwner, EventDateTime, EventDetails) " & _
			"VALUES " & _
			"('" & EventType & "','" & EventCategory & "','" & _
			CMWT_USERNAME() & "','" & NOW & "','" & Comment & "')"
		if c = "" then
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.ConnectionTimeOut = 5
			conn.Open Application("DSN_CMWT")
			conn.Execute query
			conn.Close
		else
			c.Execute query
		end if
	end if
end sub

'----------------------------------------------------------------
' function-name: CMWT_CM_CLIENTCOUNT
' function-desc: 
'----------------------------------------------------------------

Function CMWT_CM_CLIENTCOUNT ()
	Dim query, conn, cmd, rs, result
	query = "SELECT COUNT(*) AS Computers FROM (SELECT DISTINCT ResourceID FROM dbo.v_R_System) AS T1"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
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
		result = rs.Fields("Computers").value
	Else
		result = 0
	End If
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CM_CLIENTCOUNT = result
End Function

'----------------------------------------------------------------
' function-name: CMWT_CM_APPCOUNT
' function-desc: 
'----------------------------------------------------------------

Function CMWT_CM_APPCOUNT ()
	Dim query, conn, cmd, rs, result
	query = "SELECT COUNT(*) AS Apps FROM (SELECT DISTINCT DisplayName0 FROM dbo.v_GS_ADD_REMOVE_PROGRAMS) AS T1"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.ConnectionTimeOut = 5
	conn.Open Application("DSN_CMDB")
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
		result = rs.Fields("Apps").value
	Else
		result = 0
	End If
	rs.Close
	conn.Close
	Set rs = Nothing
	Set cmd = Nothing
	Set conn = Nothing
	CMWT_CM_APPCOUNT = result
End Function

'----------------------------------------------------------------
' sub-name: CMWT_WAIT
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_WAIT (intSeconds)
	startTime = Time()
	Do Until DateDiff("s", startTime, Time(), 0, 0) > intSeconds
	Loop
End Sub

'----------------------------------------------------------------
' sub-name: CMWT_LIST_SITELOGS
' sub-desc: 
'----------------------------------------------------------------

Sub CMWT_LIST_SITELOGS (CurrentFilename, ListSize)
	Dim q, objFSO, objFolder, logPath, installDir, conn, rs, rsFiles
	Dim fileName, objFile
	CMWT_DB_OPEN Application("DSN_CMDB")
	q = "SELECT TOP 1 InstallDir FROM dbo.v_Site"
	CMWT_DB_QUERY Application("DSN_CMDB"), q
	installDir = rs.Fields("InstallDir").value
	conn.Close
	rs.Close
	logPath = installDir & "\logs"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(logPath)
	Set rsFiles = CreateObject("ADODB.RecordSet")
	rsFiles.CursorLocation = adUseClient
	rsFiles.Fields.Append "filename", adVarChar, 50
	rsFiles.Open
	For each objFile in objFolder.Files 
		fileName = objFile.Name
		rsFiles.AddNew
		rsFiles.Fields("filename").value = fileName
		rsFiles.Update
	Next
	rsFiles.Sort = "filename"
	rsFiles.MoveFirst
	Response.Write "<select name=""list1"" id=""list1"" size=""" & ListSize & """ " & _
		"class=""w400 pad5"" " & _
		"onChange=""if (this.options[this.selectedIndex].value != 'null') { window.open(this.options[this.selectedIndex].value,'_top') }"">"

	Do Until rsFiles.EOF
		fileName  = rsFiles.Fields("filename").value
		If CurrentFilename <> "" And Lcase(CurrentFilename) = Lcase(fileName) Then
			Response.Write "<option value="""" selected>" & fileName & "</option>"
		Else
			Response.Write "<option value=""logview.asp?p=" & logPath & "&f=" & fileName & """>" & fileName & "</option>"
		End If
		rsFiles.MoveNext
	Loop
	Response.Write "</select>"
	rsFiles.Close
	Set rsFiles = Nothing 
End Sub

'----------------------------------------------------------------
' function-name: CMWT_WordCase
' function-desc: 
'----------------------------------------------------------------

Function CMWT_WordCase (strVal)
	If CMWT_NotNullString(strVal) Then
		If Len(strVal) > 1 Then
			CMWT_WordCase = UCase(Left(strVal,1)) & Lcase(Mid(strVal,2))
		Else
			CMWT_WordCase = Ucase(strVal)
		End If
	Else
		CMWT_WordCase = ""
	End If
End Function 

Function CMWT_DB_OfflineSort (sourceRS, fieldnames, sortName)
	Dim fn
	Set rs = CreateObject("ADODB.RecordSet")
	rs.CursorLocation = adUseClient
	For each fn in Split(fieldnames, ",")
		rs.Fields.Append fn, adVarChar, 255
	Next
	rs.Open
	Do Until sourceRS.EOF
		For each fn in Split(fieldnames, ",")
			rs.AddNew
			rs.Fields(fn).value = sourceRS.Fields(fn).value
			rs.Update
		Next
		sourceRS.MoveNext
	Loop
	rs.Sort = sortName
	rs.MoveFirst
End Function
%>
