<%
'-----------------------------------------------------------------------------
' filename....... _sm.asp
' lastupdate..... 12/09/2016
' description.... sidebar menu panel
'-----------------------------------------------------------------------------

if PageBackLink <> "" and PageBackName <> "" then
	backtab = "<a href=""" & PageBackLink & """ title=""" & PageBackName & """>" & PageBackName & "</a>:&nbsp;"
else
	backtab = ""
end if
%>
<div id="mySidenav" class="sidenav">
  <a href="javascript:void(0)" class="closebtn" onclick="closeNav()">&times;</a>
  <a href="./" title="Home">Home</a>
  <% If CMWT_ADMIN() Then %>
  <a href="admin.asp" title="Administration">Administration</a>
  <% End If %>
  <a href="./cmsite.asp" title="Site">Site</a>
  <a href="./assets.asp" title="Assets">Assets</a>
  <a href="./software.asp" title="Software">Software</a>
  <a href="./reports.asp" title="Reports">Reports</a>
  <a href="./adtools.asp" title="AD Tools">AD Tools</a>
  <a href="https://github.com/Skatterbrainz/cmwt/wiki" target="_blank" title="Help">Help</a>
  <a href="about.asp" title="About">About</a>
</div>

<span style="font-size:30px;cursor:pointer" onclick="openNav()" title="Menu">&#9776; </span>
<span style="font-size:30px;color:#00995c"><%=backtab%><%=PageTitle%></span>

<script>
function openNav() {
    document.getElementById("mySidenav").style.width = "250px";
}

function closeNav() {
    document.getElementById("mySidenav").style.width = "0";
}
</script>
