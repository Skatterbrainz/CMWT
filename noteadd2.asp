<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... noteadd2.asp
' lastupdate..... 03/20/2016
' description.... add new custom note
'-----------------------------------------------------------------------------
Response.Expires = -1
ItemID   = CMWT_GET("id", "")
ItemType = CMWT_GET("type", "")
NoteText = CMWT_GET("comm", "")

CMWT_VALIDATE ItemID, "Item Name or ID was not provided"
CMWT_VALIDATE ItemType, "Item Class or Type was not specified"
CMWT_VALIDATE NoteText, "Note Comment was not provided"

query = "INSERT INTO dbo.Notes " & _
	"(AttachedTo, AttachClass, Comment, CreatedBy, DateCreated) " & _
	"VALUES (" & _
	"'" & ItemID & "','" & Ucase(ItemType) & "','" & Replace(NoteText,"'","''") & _
	"','" & CMWT_USERNAME() & "','" & NOW & "')"
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

Caption = "Adding Note Record"
PageTitle = "ConfigMgr Web Tools"
CMWT_PageRedirect TargetURL, 1
'-----------------------------------------------------------------------------
%>
