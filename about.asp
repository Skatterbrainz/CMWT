<!-- #include file=_core.asp -->
<%
'-----------------------------------------------------------------------------
' filename....... about.asp
' lastupdate..... 03/01/2017
' description.... CMWT About page
'-----------------------------------------------------------------------------
time1 = Timer
PageTitle = "About CMWT"

CMWT_NewPage "", "", ""

%>
<!-- #include file="_sm.asp" -->
<!-- #include file="_banner.asp" -->

<table class="tf800">
	<tr>
		<td class="td6 v10" colspan="2">
			<p>Designed and developed by David Stein.  For questions or feedback, please email 
			<a href="mailto:<%=MailBox%>?subject=CMWT Feedback"><%=Application("CMWT_SupportMail")%></a>.  Thank you!</p>
			<input type="button" name="bb2" id="bb2" value="View License" class="btx w180 h30" onClick="javascript:window.open('LICENSE.TXT','licwin','width=800,height=400');" />
			<input type="button" name="bb2" id="bb2" value="Get Updates" class="btx w180 h30" onClick="javascript:window.open('https://github.com/skatterbrainz/cmwt/wiki/','gitwin','width=1200,height=600,toolbar=yes,scrollbars=yes,resizable=yes,titlebar=yes,status=yes,menubar=yes');" />
		</td>
	</tr>
	<%
	For each sv in Split("CMWT_TITLE,CMWT_SUBTITLE,CMWT_VERSION,CMWT_BUILD,CM_SITECODE,CMWT_DOMAINSUFFIX",",")
		Response.Write "<tr>" & _
			"<td class=""td6 v10 bgGray w200"">" & sv & "</td>" & _
			"<td class=""td6 v10 right bgBlue"">" & Application(sv) & "</td>" & _
			"</tr>"
	Next
	%>
</table>
	
<br/>
<div class="tf800"><h3>Session Data</h3></div>

<table class="tf800">
	<%
	For each sv in Session.Contents
		Response.Write "<tr class=""tr1"">" & _
			"<td class=""td6 v10 bgGray w200"">" & sv & "</td>" & _
			"<td class=""td6 v10 bgBlue"">" & Session(sv) & "</td>" & _
			"</tr>"
	Next
	Response.Write "<tr class=""tr1"">" & _
		"<td class=""td6 v10 bgGray w200"">BROWSER TYPE</td>" & _
		"<td class=""td6 v10 bgBlue"">" & CMWT_BROWSER_TYPE() & "</td>" & _
		"</tr>"
	%>
</table>
	
<% CMWT_Footer() %>

</body>
</html>