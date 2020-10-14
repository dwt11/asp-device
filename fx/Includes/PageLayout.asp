<%
'This page just contains functions to render page layout.
Function render_pageHeader()
	'This function renders the page header. It includes headers too.
	%>
	<table width="960" align="center" cellpadding="0" cellspacing="0" border="0" background="images/PageBg.jpg">
		<tr height="70">
			<td width="33">&nbsp;	
			
			</td>
			<td align="left" valign="bottom">
				<a href="http://www.InfoSoftGlobal.com" target="_blank"><img src="Images/IGPLogo.jpg" alt="InfoSoft Global" border="0"></a>
			</td>
			<td align="right" valign="bottom">
				<img src="Images/TopRightText.gif" border="0">	
			</td>
			<td width="37">&nbsp;	
			
			</td>
		</tr>

		<tr>
			<td width="33">		
			</td>
			<td height="1" colspan="2" bgColor="#EEEEEE">
			</td>
			<td width="37">		
			</td>
		</tr>
	<%
End Function

Function render_pageTableOpen()
	'This function renders the page table open
	%>
	<tr>
		<td height="10" colspan="4">
		</td>
	</tr>

	<tr>
		<td width="33">	
		</td>
		<td colspan="2">		
	<%
End Function

Function render_pageTableClose()
	'This function renders the page table closing tags
	%>
			<br>
			</td>
			<td width="37">&nbsp;
			
			</td>
		</tr>	
		
		<tr>
			<td width="33">		
			</td>
			<td height="1" colspan="2" bgColor="#EEEEEE">
			</td>
			<td width="37">		
			</td>
		</tr>
		
		<tr>
			<td height="4" colspan="4">		
			</td>			
		</tr>
		
		<tr>
			<td width="33">		
			</td>
			<td colspan="2" align="center">
			<span class="text">This application was built using <a href="http://www.InfoSoftGlobal.com/FusionCharts" target="_blank"><b>FusionCharts v3</b></a> - &quot;Animated flash charts for the web&quot;.</span>
			</td>
			<td width="33">
			</td>
		</tr>
		
		<tr>
			<td width="33">		
			</td>
			<td colspan="2" align="center">
			<span class="text">?All Rights Reserved</span>
			</td>
			<td width="33">
			</td>
		</tr>
		
		<tr>
			<td height="4" colspan="4">		
			</td>			
		</tr>
	</table>
	<%
End Function

'This function draws a separator line between two tables
Function drawSepLine()
%>
	<table width="875">
		<tr>
			<td width="33">		
			</td>
			<td height="1" colspan="2" bgColor="#EEEEEE">
			</td>
			<td width="37">		
			</td>
		</tr>
	</table>
<%
End Function

'This function renders the year selection form
Function render_yearSelectionFrm(action)
%>
<!-- Code to render the form for year selection and animation selection -->
<tr>
	<td width="33">		
	</td>
	<td height="1" colspan="2" bgColor="#FFFFFF">
	</td>
	<td width="37">		
	</td>
</tr>

<form name="frmYr" action="sb_fx.asp?action=<%=action%>" method="post" id="frmYr">
<tr height="30">
	<td width="33">		
	</td>
	<td height="22" colspan="2" align="center" bgColor="#EEEEEE" valign="middle">
	<nobr>
	<span class="textbolddark">选择年: </span>
	<%	
	'Retrieve the years

	Set oRs = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT DISTINCT YEAR(jx_date) As Year22 FROM sbjx"
	oRs.Open strSQL, Conn
	'Render them in drop down box	
	While not oRs.EOF
		if int(intYear) = int(oRs("Year22")) then
			Response.Write("<input type='radio' name='Year' value='") & ors("Year22") & ("' checked><span class='text'>") & ors("Year22") & "</span>&nbsp;&nbsp;"
		else
			Response.Write("<input type='radio' name='Year' value='") & ors("Year22") & ("'><span class='text'>") & ors("Year22") & "</span>&nbsp;&nbsp;"
		end if		 
		oRs.MoveNext()
	Wend
	%>		
	<span class="textbolddark"><span class='text'>&nbsp;&nbsp;&nbsp;</span>动画: </span>
	<%
		if getAnimationState()="1" then
	%>	
	<input type="radio" name="animate" value="1" checked><span class="text">开</span>
	<input type="radio" name="animate" value="0"><span class="text">关</span>
	<%
		else		
	%>
	<input type="radio" name="animate" value="1"><span class="text">开</span>
	<input type="radio" name="animate" value="0" checked><span class="text">关</span>
	<%
		end if
	%>
	<span class='text'>&nbsp;&nbsp;</span>
	<input type="submit" class="button" value="Go" id="submit"  name="submit" 1>
	
	</nobr>	
	</td>
	<td width="37">		
	</td>
</tr>
</form>	

<tr>
	<td width="33">		
	</td>
	<td height="1" colspan="2" bgColor="#FFFFFF">
	</td>
	<td width="37">		
	</td>
</tr>

<tr>
	<td width="33">		
	</td>
	<td height="1" colspan="2" bgColor="#EEEEEE">
	</td>
	<td width="37">		
	</td>
</tr>
<!-- End code to render form -->
<%
End Function
%>