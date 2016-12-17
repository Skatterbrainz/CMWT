<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... collmem2.asp
' lastupdate..... 12/15/2016
' description.... collection direct-rule membership tools
'-----------------------------------------------------------------------------
time1 = Timer

CollectionID1 = CMWT_GET("cid1", "")
CollectionID2 = CMWT_GET("cid2", "")
ActionType1   = CMWT_GET("a1", "")
ActionType2   = CMWT_GET("a2", "")
MemberList1   = CMWT_GET("m1", "")
MemberList2   = CMWT_GET("m2", "")
WhatIf        = CMWT_GET("w", "")
WaitTime = 2

If ActionType1 = "" And ActionType2 = "" Then
	CMWT_STOP "Action (Copy or Move) was not selected for either list"
End If

If MemberList1 = "" And MemberList2 = "" Then
	CMWT_STOP "No Collection members were selected"
End If

MemberList1 = Replace(MemberList1, " ", "")
MemberList2 = Replace(MemberList2, " ", "")

'Response.Write "<br/>m1: " & MemberList1
'Response.Write "<br/>m2: " & MemberList2
'response.end

Dim conn, cmd, rs, rows
rows = 0

Sub CMWT_CM_CopyMoveMembers (Collection1, Collection2, MachineList, Mode)
	Dim caption, try, m
	Select Case Ucase(Mode)
		Case "COPY":
			CMWT_DB_OPEN Application("DSN_CMDB")
			For each m in Split(MachineList, ",")
				If CMWT_CM_IsCollectionMember(conn, m, Collection2) = True Then
					caption = "Warning: Resource " & m & " is already a member of collection " & Collection2
				Else
					try = CMWT_Add_CmCollectionMember (Collection2, m)
					'try = CMWT_CM_AddCollectionMember (Collection2, m)
					If try = 0 Then
						caption = "Success: Added Resource " & m & " to Collection " & Collection2
						CMWT_LogEvent "", "INFO", "CM COLLECTION COPY", m & " was added to collection: " & Collection2
					Else
						caption = "Error: Failed to add Resource " & m & " to Collection " & Collection2
						CMWT_LogEvent "", "ERROR", "CM COLLECTION COPY", m & " was not added to collection: " & Collection2 & " (" & try & ")"
					End If
				End If
				rows = rows + 1
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10"">" & m & "</td>" & _
					"<td class=""td6 v10"">" & Mode & "</td>" & _
					"<td class=""td6 v10"">" & caption & "</td></tr>"
				CMWT_WAIT(WaitTime)
			Next
			CMWT_DB_CLOSE()
		Case "MOVE":
			CMWT_DB_OPEN Application("DSN_CMDB")
			For each m in Split(MachineList, ",")
				If CMWT_CM_IsCollectionMember(conn, m, Collection2) = True Then
					caption = "Warning: Already a member of this collection"
				Else
					try = CMWT_Add_CmCollectionMember (Collection2, m)
					'try = CMWT_CM_AddCollectionMember (Collection2, m)
					If try = 0 Then
						caption = "Success: Added Resource " & m & " to Collection " & Collection2
						CMWT_LogEvent "", "INFO", "CM COLLECTION COPY", m & " was added to collection: " & Collection2
						try = CMWT_Remove_CmCollectionMember (m, Collection1)
						'try = CMWT_CM_RemoveCollectionMember (Collection1, m)
						If try = 0 Then
							caption = "Success: Removed Resource " & m & " from Collection " & Collection1
							CMWT_LogEvent "", "INFO", "CM COLLECTION REMOVE", m & " was removed from collection: " & Collection1
						Else
							caption = "Error: Failed to remove Resource " & m & " from Collection " & Collection1
							CMWT_LogEvent "", "ERROR", "CM COLLECTION REMOVE", m & " was not removed from collection: " & Collection1 & " (" & try & ")"
							If eventList <> "" Then
								eventList = eventList & "|ERROR,CM COLLECTION REMOVE," & m & " was not removed from collection: " & Collection1
							Else
								eventList = "ERROR,CM COLLECTION REMOVE," & m & " was not removed from collection: " & Collection1
							End If
						End If
					Else
						caption = "Error: Failed to add Resource " & m & " to Collection " & Collection2
						CMWT_LogEvent "", "ERROR", "CM COLLECTION COPY", m & " was not added to collection: " & Collection2 & " (" & try & ")"
					End If
				End If
				rows = rows + 1
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10"">" & m & "</td>" & _
					"<td class=""td6 v10"">" & Mode & "</td>" & _
					"<td class=""td6 v10"">" & caption & "</td></tr>"
				CMWT_WAIT(WaitTime)

			Next
			CMWT_DB_CLOSE()
		Case "REMOVE"
			CMWT_DB_OPEN Application("DSN_CMDB")
			For each m in Split(MachineList, ",")
				'try = CMWT_CM_RemoveCollectionMember (Collection1, m)
				try = CMWT_Remove_CmCollectionMember (m, Collection1)
				If try = 0 Then
					caption = "Success: Removed Resource from Collection"
					CMWT_LogEvent "", "INFO", "CM COLLECTION REMOVE", m & " was removed from collection: " & Collection1
				Else
					caption = "Error: Failed to remove Resource from Collection"
					CMWT_LogEvent "", "ERROR", "CM COLLECTION REMOVE", m & " was not removed from collection: " & Collection1 & " (" & try & ")"
					If eventList <> "" Then
						eventList = eventList & "|ERROR,CM COLLECTION REMOVE," & m & " was not removed from collection: " & Collection1
					Else
						eventList = "ERROR,CM COLLECTION REMOVE," & m & " was not removed from collection: " & Collection1
					End If
				End If
				rows = rows + 1
				Response.Write "<tr class=""tr1"">" & _
					"<td class=""td6 v10"">" & m & "</td>" & _
					"<td class=""td6 v10"">" & Mode & "</td>" & _
					"<td class=""td6 v10"">" & caption & "</td></tr>"
				CMWT_WAIT(WaitTime)

			Next
			CMWT_DB_CLOSE()
		
	End Select
End Sub

PageTitle    = "Collection Tools"
PageBackLink = "collmem.asp"
PageBackName = "Collection Tools"

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

Response.Write "<table class=""tfx"">" & _
	"<tr>" & _
	"<td class=""td6 v10 bgGray"">Computer</td>" & _
	"<td class=""td6 v10 bgGray"">Operation</td>" & _
	"<td class=""td6 v10 bgGray"">Result</td></tr>"

If MemberList1 <> "" And ActionType1 <> "" And CollectionID2 <> "" Then 
	SourceCollection = CollectionID1
	TargetCollection = CollectionID2
	CMWT_CM_CopyMoveMembers SourceCollection, TargetCollection, MemberList1, ActionType1
End If

If MemberList2 <> "" And ActionType2 <> "" And CollectionID1 <> "" Then
	SourceCollection = CollectionID2
	TargetCollection = CollectionID1
	CMWT_CM_CopyMoveMembers SourceCollection, TargetCollection, MemberList2, ActionType2
End If

Response.Write "<tr><td class=""td6 v10 bgGray"" colspan=""3"">" & rows & " requests processed</td></tr></table>"

CMWT_FOOTER()
Response.Write "</body></table>"
%>
