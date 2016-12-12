<%
On Error Resume Next
Set objshell = Server.CreateObject("Wscript.Shell")
'response.write objShell.ExpandEnvironmentStrings("%USERNAME%")
'cmdStatement = "powershell.exe -ExecutionPolicy ByPass -Command Stop-Service MpsSvc -Force"
'exitCode = objShell.Run( cmdStatement, 0, True )
'response.write exitCode
x = objShell.run("powershell -executionpolicy bypass -windowstyle hidden -file f:\cmwt\scripts\test.ps1")
response.write x
%>