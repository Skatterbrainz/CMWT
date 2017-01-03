<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... logins.asp
' lastupdate..... 01/02/2017
' description.... device logins
'-----------------------------------------------------------------------------
time1 = Timer

SortBy   = CMWT_GET("s", "ComputerName")
FilterFN = CMWT_GET("fn", "")
FilterFV = CMWT_GET("fv", "")
QueryON  = CMWT_GET("qq", "")

if FilterFV <> "" Then
	PageTitle    = "Device Logins: " & FilterFV
else
	PageTitle    = "Device Logins"
end if
PageBackLink = "reports.asp"
PageBackName = "Reports"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%

query = "SELECT * FROM (SELECT " & _
	"MachineResourceName AS ComputerName, " & _
	"MachineResourceID AS ResourceID, " & _
	"UniqueUserName AS UserID, " & _
	"NumberOfLogins AS Logins, " & _
	"LastLoginTime AS LastLogin, " & _
	"ConsoleMinutes AS LoginMins " & _
	"FROM dbo.v_UserMachineIntelligence) AS T1 "
If FilterFN <> "" Then
	query = query & _
		" WHERE (T1." & FilterFN & " = '" & FilterFV & "') " & _
		"ORDER BY T1." & SortBy
	filtered = True
Else
	query = query & " ORDER BY " & SortBy	
End If

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TableGridFilter rs, "", "logins.asp", "", "", "logins.asp"
'CMWT_DB_TABLEGRID rs, "", "logins.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY() 
CMWT_FOOTER()

Response.Write "</body></html>"
%>
