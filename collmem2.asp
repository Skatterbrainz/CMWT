<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... collmem2.asp
' lastupdate..... 12/08/2016
' description.... collection direct-rule membership tools
'-----------------------------------------------------------------------------
time1 = Timer

CollectionID1 = CMWT_GET("cid1", "")
CollectionID2 = CMWT_GET("cid2", "")
ActionType1   = CMWT_GET("a1", "")
ActionType2   = CMWT_GET("a2", "")
MemberList1   = CMWT_GET("m1", "")
MemberList2   = CMWT_GET("m2", "")

'Response.Write "<br/>cid1: " & CollectionID1
'Response.Write "<br/>cid2: " & CollectionID2
Response.Write "<br/>action1: " & ActionType1 & " to: " & CollectionID1
Response.Write "<br/>action2: " & ActionType2 & " to: " & CollectionID2
Response.Write "<br/>m1: " & MemberList1
Response.Write "<br/>m2: " & MemberList2
%>