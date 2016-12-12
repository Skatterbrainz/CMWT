<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... noteedit2.asp
' lastupdate..... 12/04/2016
' description.... update and existing custom note
'-----------------------------------------------------------------------------
Response.Expires = -1
NoteID   = CMWT_GET("id", "")
ItemID   = CMWT_GET("iid", "")
ItemType = CMWT_GET("type", "")
NoteText = CMWT_GET("comm", "")

CMWT_VALIDATE ItemID, "Item Name or ID was not provided"
CMWT_VALIDATE NoteID, "Note record ID was not provided"
CMWT_VALIDATE ItemType, "Item Class or Type was not specified"
CMWT_VALIDATE NoteText, "Note Comment was not provided"

query = "UPDATE dbo.Notes " & _
	"SET Comment='" & Replace(NoteText, "'", "''") & "'," & _
	"CreatedBy='" & CMWT_USERNAME() & "'," & _
	"DateCreated='" & NOW & "' " & _
	"WHERE NoteID=" & NoteID

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

targetURL = CMWT_PageLink (ItemType, ItemID)

Caption   = "Updating Note Record"
PageTitle = "ConfigMgr Web Tools"

CMWT_PageRedirect TargetURL, 1
%>