<!-- #include file="./_core.asp" -->
<%
'****************************************************************
' Filename..: reload.asp
' Date......: 03/20/2016
' Purpose...: reload session state
'****************************************************************

'Application.Lock
'Application.Contents.RemoveAll()
'Application.Unlock
'Session.Contents.RemoveAll()
'Session.Abandon()

caption = "Reloading Application"

CMWT_PageRedirect "./", 2

%>
