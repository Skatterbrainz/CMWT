<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... dupefiles.asp
' lastupdate..... 12/05/2016
' description.... duplicate software files on specified client
'-----------------------------------------------------------------------------
time1 = Timer

cn = CMWT_GET("cn", "")
fn = CMWT_GET("fn", "")
CMWT_VALIDATE cn, "Computer Name was not specified"
CMWT_VALIDATE fn, "Filename was not specified"

QueryOn = CMWT_GET("qq", "")
SortBy  = CMWT_GET("s","FileName")

PageTitle    = "Duplicate Files"
PageBackLink = "device.asp?cn=" & cn
PageBackName = "Devices: " & cn

CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
Response.Write "<table class=""tfx""><tr><td class=""pad6 v10 bgDarkGray"">" & _
	"Instances of file [" & fn & "] inventoried on the specified computer</td></tr></table>"

query = "SELECT DISTINCT " & _
		"dbo.v_GS_SoftwareFile.FileName, " & _
		"dbo.v_GS_SoftwareFile.FilePath, " & _
		"dbo.v_GS_SoftwareFile.FileSize, " & _
		"dbo.v_GS_SoftwareFile.FileDescription, " & _
		"dbo.v_GS_SoftwareFile.FileVersion, " & _
		"dbo.v_GS_SoftwareFile.FileModifiedDate " & _
	"FROM  " & _
		"dbo.v_R_System INNER JOIN " & _
		"dbo.v_GS_SoftwareFile ON dbo.v_R_System.ResourceID = dbo.v_GS_SoftwareFile.ResourceID " & _
	"WHERE " & _
		"(dbo.v_R_System.Name0 = '" & cn & "') " & _
		"AND " & _
		"(dbo.v_GS_SoftwareFile.FileName='" & fn & "') " & _
	"ORDER BY " & SortBy

Dim conn, cmd, rs
CMWT_DB_QUERY Application("DSN_CMDB"), query
CMWT_DB_TABLEGRID rs, "", "dupefiles.asp", ""
CMWT_DB_CLOSE()
CMWT_SHOW_QUERY()
CMWT_Footer()
Response.Write "</body></html>"
%>