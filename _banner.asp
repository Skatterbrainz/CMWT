	<table class="tfx bgMedBlue">
		<tr>
			<td class="v10 pad6 cGray">
			CMWT - Configuration Manager Web Tools :: Today is <%=FormatDateTime(Now,vbLongDate)%>
			</td>
			<td class="v10 pad6 right w500">
				<form name="formsearch" id="formsearch" method="post" action="search.asp">
					<input type="button" name="bprint" id="bprint" class="btx w120 h28" value="Print" onClick="javascript:print();" title="Print Page" /> 
					<input type="text" name="q" id="q" class="pad5 v10 w180" maxlength="50" title="Enter Search Value" />
					<input type="submit" name="s1" id="s1" class="btx w40 h28" value="..." title="Search" />
				</form>
			</td>
		</tr>
	</table>
	