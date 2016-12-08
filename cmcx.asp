<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... cmcx.asp
' lastupdate..... 12/04/2016
' description.... add or remove resources to/from a collection
'-----------------------------------------------------------------------------
Response.Expires = -1

cnx = CMWT_GET("cn", "")
mx  = CMWT_GET("mx", "")
cid = CMWT_GET("cid", "")
zz  = CMWT_GET("z", "device")

PageTitle = "Collection Membership"

CMWT_VALIDATE cnx, "Resource Names were not provided"
CMWT_VALIDATE mx, "Operation Type parameter was not provided"
CMWT_VALIDATE cid, "Collection ID was not provided"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx""><tr>" & _
	"<td class=""td6 v10 bgGray w150"">Resource</td>" & _
	"<td class=""td6 v10 bgGray w150"">Action</td>" & _
	"<td class=""td6 v10 bgGray"">Result</td></tr>"

Dim conn, cmd, rs

Select Case Ucase(mx)
	Case "ADD":
		CMWT_DB_OPEN Application("DSN_CMDB")
		For each cnn in Split(cnx, ",")
			If CMWT_CM_IsCollectionMember(conn, cnn, cid) = True Then
				caption = "Already a member of this collection"
			Else
				try = CMWT_CM_AddCollectionMember (cid, cnn)
				If try = 1 Then
					caption = "Added Resource to Collection"
					CMWT_LogEvent "", "INFO", "CM COLLECTION ADD", cnn & " was added to collection: " & cid
				Else
					caption = "ERROR: Failed to add Resource to Collection"
					CMWT_LogEvent "", "ERROR", "CM COLLECTION ADD", cnn & " was not added to collection: " & cid & " (" & try & ")"
				End If
			End If
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10"">" & cnn & "</td>" & _
				"<td class=""td6 v10"">" & mx & "</td>" & _
				"<td class=""td6 v10"">" & caption & "</td></tr>"
			CMWT_WAIT(2)
		Next
		CMWT_DB_CLOSE()
	Case "REM","REMOVE":
		CMWT_DB_OPEN Application("DSN_CMDB")
		For each cnn in Split(cnx, ",")
			try = CMWT_CM_RemoveCollectionMember (cid, cnn)
			If try = 1 Then
				caption = "Removed Resource from Collection"
				CMWT_LogEvent "", "INFO", "CM COLLECTION REMOVE", cnn & " was removed from collection: " & cid
			Else
				caption = "ERROR: Failed to remove Resource from Collection"
				CMWT_LogEvent "", "ERROR", "CM COLLECTION REMOVE", cnn & " was not removed from collection: " & cid & " (" & try & ")"
				If eventList <> "" Then
					eventList = eventList & "|ERROR,CM COLLECTION REMOVE," & cn & " was not removed from collection: " & cid
				Else
					eventList = "ERROR,CM COLLECTION REMOVE," & cn & " was not removed from collection: " & cid
				End If
			End If
			Response.Write "<tr class=""tr1"">" & _
				"<td class=""td6 v10"">" & cnn & "</td>" & _
				"<td class=""td6 v10"">" & mx & "</td>" & _
				"<td class=""td6 v10"">" & caption & "</td></tr>"
			CMWT_WAIT 2
		Next
		CMWT_DB_CLOSE()
	Case Else:
		CMWT_STOP "invalid operation code requested."
End Select

Response.Write "</table>"

Response.Write "<p class=""tf800"">Note: Actual change may take a few seconds to appear in the CMWT console.</p>"

targetURL = "collection.asp?id=" & cid & "&ks=2"

Response.Write "<p><input type=""button"" name=""b1"" id=""b1"" " & _
	"class=""btx w150 h32"" value=""Continue"" onClick=""document.location.href='" & TargetURL & "'"" /></p>"
CMWT_Footer()

'CMWT_PageRedirect TargetURL, 8
%>