<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... sitelogs.asp
' lastupdate..... 12/07/2016
' description.... home page
'-----------------------------------------------------------------------------
time1 = Timer
SortBy = CMWT_GET("s", "FileName")
PageTitle    = "Site Server Logs"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"

SelfLink   = "sitelogs.asp"
SortLink   = "sitelogs.asp"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<%
Dim conn, cmd, rs
CMWT_DB_OPEN Application("DSN_CMDB")
q = "SELECT TOP 1 InstallDir FROM dbo.v_Site"
CMWT_DB_QUERY Application("DSN_CMDB"), q
Response.Write "<table class=""t1x""><tr>"
installDir = rs.Fields("InstallDir").value
CMWT_DB_CLOSE()

logPath = installDir & "\logs"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(logPath)
Set rsFiles = CreateObject("ADODB.RecordSet")
rsFiles.CursorLocation = adUseClient
rsFiles.Fields.Append "filename", adVarChar, 50
rsFiles.Fields.Append "datecreated", adVarChar, 50
rsFiles.Fields.Append "datemodified", adVarChar, 50
rsFiles.Fields.Append "filesize", adVarChar, 50
rsFiles.Open

For each objFile in objFolder.Files 
	fileName = objFile.Name
	filePath = logPath & "\" & fileName
	fileDate1 = objFile.DateCreated
	fileDate2 = objFile.DateLastModified
	fileSize  = objFile.Size
	rsFiles.AddNew
	rsFiles.Fields("filename").value = fileName
	rsFiles.Fields("datecreated").value = fileDate1
	rsFiles.Fields("datemodified").value = fileDate2
	rsFiles.Fields("filesize").value = fileSize
	rsFiles.Update
Next
rsFiles.Sort = SortBy
rsFiles.MoveFirst

Response.Write "<table class=""tfx""><tr>"
For each fn in Split("FileName,DateCreated,DateModified,FileSize",",")
	Response.Write "<td class=""td6 v10 bgGray"">" & CMWT_SORTLINK(SortLink, fn, SortBy) & "</td>"
Next
Response.Write "</tr>"
Do Until rsFiles.EOF
	fileName  = rsFiles.Fields("filename").value
	fileDate1 = rsFiles.Fields("datecreated").value
	fileDate2 = rsFiles.Fields("datemodified").value
	fileSize  = rsFiles.Fields("filesize").value
	fileLink = "<a href=""logview.asp?p=" & logPath & "&f=" & fileName & """ target=""_blank"" title=""Open Log File"">" & fileName & "</a>"
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""td6 v10"">" & fileLink & "</td>" & _
		"<td class=""td6 v10"">" & fileDate1 & "</td>" & _
		"<td class=""td6 v10"">" & fileDate2 & "</td>" & _
		"<td class=""td6 v10"">" & fileSize & "</td>" & _
		"</tr>"
	rsFiles.MoveNext
Loop
Response.Write "</table>"

rsFiles.Close
Set rsFiles = Nothing 

CMWT_FOOTER()
Response.WRite "</body></html>"
%>
