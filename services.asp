<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... services.asp
' lastupdate..... 05/23/2017
' description.... site system windows services status
'-----------------------------------------------------------------------------
time1 = Timer
SortBy  = CMWT_GET("s", "DisplayName")
QueryON = CMWT_GET("qq", "")

PageTitle    = "Services"
PageBackLink = "cmsite.asp"
PageBackName = "Site Hierarchy"
CMWT_NewPage "", "", ""
%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->
<%
wmi_columns = "DisplayName,Name,StartMode,State,StartName"
wmi_class   = "Win32_Service"
query = "SELECT " & wmi_columns & " FROM " & wmi_class
CMWT_WMI_TABLEGRID ".", wmi_columns, wmi_class, "", sortby, "Name=service.asp?sn="
CMWT_SHOW_QUERY() 
CMWT_Footer()
Response.Write "</body></html>"
%>
