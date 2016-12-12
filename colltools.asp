<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... colltools.asp
' lastupdate..... 12/12/2016
' description.... collection tools
'-----------------------------------------------------------------------------
time1 = Timer

PageTitle = "Collection Tools"
CollID    = CMWT_GET("cid", "")
GroupNum  = CMWT_GET("group", "")
ActType   = CMWT_GET("atyp", "")
ActionID1 = CMWT_GET("xx1", "")
ActionID2 = CMWT_GET("xx2", "")
ActionID3 = CMWT_GET("xx3", "")
QueryOn   = CMWT_GET("qq", "")

CMWT_VALIDATE CollID, "Collection ID was not specified"
CMWT_VALIDATE GroupNum, "Collection Tools action group number was not provided"

Select Case GroupNum
	Case "1"
		actName = "Client-Actions"
		actCode = ActionID1
		Comment = ""
		CommandString = "Send-ClientAction.ps1 -Collection " & CollID & " -Action " & actCode
	
	Case "2"
		actName = "Client-Tools"
		actCode = ActionID2
		Comment = ""
		CommandString = "Send-ClientTool.ps1 -Collection " & CollID & " -Action " & actCode
	
End Select

query = "INSERT INTO dbo.Tasks " & _
	"(ActivityName,ActivityType,CreatedBy,DateTimeCreated,Comment,CommandString) " & _
	"VALUES " & _
	"('" & actName & "','" & actType & "','" & CMWT_USERNAME() & "','" & NOW & "','" & Comment & "','" & CommandString & "')"

On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeOut = 5
conn.Open Application("DSN_CMWT")
If err.Number <> 0 Then
	CMWT_STOP "database connection failure"
End If
conn.Execute query
conn.Close
Set conn = Nothing

PageTitle    = "Collection Tools"
PageBackLink = "collection.asp?id=" & CollID
PageBackName = "Collection"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Response.Write "<table class=""tfx"">" & _
	"<tr class=""h100""><td class=""td6 v10 ctr bgDarkGray"">" & _
	"<p>" & actName & " request submitted into process queue.</p>" & _
	"<p><a href=""collection.asp?id=" & CollID & """ title=""Return to Collection"">Return to Collection</a></p>" & _
	"</td></tr></table>"

Response.Write "</body></html>"
%>