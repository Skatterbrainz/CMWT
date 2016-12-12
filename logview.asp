<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... logView.asp
' lastupdate..... 12/07/2016
' description.... home page
'-----------------------------------------------------------------------------
time1 = Timer
LogPath = CMWT_GET("p", "")
LogFile = CMWT_GET("f", "")
FindVal = CMWT_GET("v", "")
SortBy  = CMWT_GET("s", "datetime desc")
CMWT_VALIDATE LogPath, "Log folder path was not specified"
CMWT_VALIDATE LogFile, "Log file name was not specified"

PageTitle  = LogFile
PageBackLink = "sitelogs.asp"
PageBackName = "Site Server Logs"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<%
InputLogFile = LogPath & "\" & LogFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(InputLogFile, ForReading)

If err.Number = 0 Then
	Response.Write "<p>Reading: " & InputLogFile & "</p>"
	strFileContents = objFile.ReadAll
	objFile.Close
Else
	Response.Write "<p>Error: " & err.Number & ": " & err.Description & "</p>"
End If
%>
<table class="tfx">
	<tr>
		<td>
			<form name="form1" id="form1" method="post" action="logview.asp?p=<%=LogPath%>&f=<%=LogFile%>">
				<input type="text" name="v" id="v" class="pad5 v10 w200" title="Filter on Value" maxlength="50" />
				<input type="submit" name="b1" id="b1" class="btx w30 h28" value="..." />
			</form>
		</td>
		<td class="w500 right">
			<% 
			'CMWT_LIST_SITELOGS LogFile, 1 
			%>
		</td>
	<tr>
</table>
<%
Sub CMWT_LIST_LogFiles ()
	Dim rsFiles
	Set rsFiles = CreateObject("ADODB.RecordSet")
	rsFiles.CursorLocation = adUseClient
	rsFiles.Fields.Append "eventinfo", adVarChar, 500
	rsFiles.Fields.Append "source", adVarChar, 255
	rsFiles.Fields.Append "datetime", adVarChar, 50
	rsFiles.Fields.Append "thread", adVarChar, 255
	rsFiles.Open

	For each strLine in Split(strFileContents, vbCRLF)
		' sample: "---->: An event has occurred, the event code is: 0x1  $$<SMS_SITE_SYSTEM_STATUS_SUMMARIZER><12-06-2016 03:11:46.039+300><thread=4840 (0x12E8)>"
		if Trim(strLine) <> "" Then
			logRowSet = Split(strLine, "$$")
			' sample: ("---->: An event has occurred, the event code is: 0x1  ", "<SMS_SITE_SYSTEM_STATUS_SUMMARIZER><12-06-2016 03:11:46.039+300><thread=4840 (0x12E8)>")
			LogHead = Trim(logRowSet(0))
			' sample: "---->: An event has occurred, the event code is: 0x1"
			If Ubound(logRowSet) > 0 then
				logTail = Split(Trim(logRowSet(1)),"><")
			Else
				logTail = Array("...","...","...")
			End If
			' sample: ("<SMS_SITE_SYSTEM_STATUS_SUMMARIZER", "12-06-2016 03:11:46.039+300", "thread=4840 (0x12E8)>"
		'	logTail = Array("1","2","3")

			txtSource = logTail(0)
			txtDate   = logTail(1)
			txtThread = logTail(2)
		else
			logHead = ". . . ."
			txtSource = "..."
			txtDate   = "..."
			txtThread = "..."
		end if
		rsFiles.AddNew
		rsFiles.Fields("eventinfo").value = LogHead
		rsFiles.Fields("source").value = txtSource
		rsFiles.Fields("datetime").value = txtDate
		rsFiles.Fields("thread").value = txtThread
		rsFiles.Update
	Next

	Response.Write "<table class=""tfx"">" & _
		"<tr><td class=""bgGray td6 v10"">Event Description</td>" & _
		"<td class=""bgGray td6 v10"">Source</td>" & _
		"<td class=""bgGray td6 v10"">Date/Time</td><td class=""td6 bgGray v10"">Thread</td></tr>"
		
	rsFiles.Sort = SortBY
	rsFiles.MoveFirst

	Do Until rsFiles.EOF
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10"">" & rsFiles.Fields("eventinfo").value & "</td>" & _
			"<td class=""td6 v10"">" & rsFiles.Fields("source").value & "</td>" & _
			"<td class=""td6 v10"">" & rsFiles.Fields("datetime").value & "</td>" & _
			"<td class=""td6 v10"">" & rsFiles.Fields("thread").value & "</td>" & _
			"</tr>"
		rsFiles.MoveNext
	Loop
	rsFiles.Close
	Set rsFiles = Nothing
	Response.Write "</table>"
End Sub

CMWT_LIST_LogFiles()

CMWT_FOOTER()
Response.WRite "</body></html>"
%>
