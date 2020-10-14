<%@language=vbscript codepage=936 %>
<%
Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->

<%
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql
if request("action")="pxjhzj" then call pxjhzj
if request("action")="" then call main
if request("action")="bb" then call bb    '用于报表输出,进行选择
'1公司报表  2分厂报表
'1培训计划  2培训总结

'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
sub main()
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>培训计划总结管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write"<br><br><br><form method='post' action='pxjhzj_bb.asp' name='form1' >" & vbCrLf
response.write "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>培训计划总结报表输出</strong></td>"& vbCrLf
response.write "  </tr> </table> "& vbCrLf
response.write "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf

response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>选择月份：</strong></td> " & vbCrLf
response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>" & vbCrLf
response.write"<input name='pxjhzj_date' type='text' value="&year(now())&"-"&month(now())&" >" & vbCrLf
response.write"<a href='#' onClick=""popUpCalendar(this,pxjhzj_date, 'yyyy-mm'); return false;"">" & vbCrLf
response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a>  只需在需要的月份选择任意一个日期即可</td></tr>"& vbCrLf
'response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>" & vbCrLf
'response.write"<strong>报表分类：</strong></td>" & vbCrLf
'response.write "<td><select name='bbclass' size='1'>" & vbCrLf
'response.write "<option value='1'>公司报表</option> " & vbCrLf
'response.write"<option value='2'>分厂报表</option>" & vbCrLf
'response.write"<select></td></tr>" & vbCrLf
'response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>" & vbCrLf
'response.write"<strong>培训分类：</strong></td>" & vbCrLf
'response.write "<td><select name='pxjhorzj' size='1'>" & vbCrLf
'response.write "<option value='1'>培训计划</option> " & vbCrLf
'response.write"<option value='2'>培训总结</option>" & vbCrLf
'response.write"<select></td></tr>" & vbCrLf
response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>" & vbCrLf
response.write"<input name='action' type='hidden' id='action' value='pxjhzj'><input  type='submit' name='Submit' value='输  出' style='cursor:hand;'></td>  </tr>" & vbCrLf
response.write"</table></form>" & vbCrLf
response.write "</body></html>"

end sub

sub pxjhzj()
'判断数据库中是否有所选月份的培训计划或总结，并调用相关输出
dim sqlbb,rsbb
	'if request("pxjhorzj")=1 then 
		    sqlbb="SELECT * from pxjh where year="&year(request("pxjhzj_date"))&" and month="&month(request("pxjhzj_date"))
		    'sqlbb="SELECT * from pxjh where year=2008 and month=3"
   			set rsbb=server.createobject("adodb.recordset")
    		rsbb.open sqlbb,conne,1,1
   			if rsbb.eof and rsbb.bof then 
				response.write"<Script Language=Javascript>window.alert('此月培训内容还没有添加');history.go(-1);</Script>"
			else
			'response.write year(request("pxjhzj_date"))&"aaaaa"&month(request("pxjhzj_date"))
			  call bb1(year(request("pxjhzj_date")),month(request("pxjhzj_date")))
			  'call bb1("2008","2")
			
			'if request("bbclass")=2 then call bb2(request("year"),request("month"))
			end if
			rsbb.close 
	'end if 		


end sub 


'bb1培训计划, bb2 培训总结
function bb1(p_year,p_month)
dim titlename
titlename="培训报表"&p_year&"年"&p_month&"月"
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&titlename&".xls"' 

	dim sqlbb,rsbb  
    sqlbb="SELECT * from pxjh where year="&p_year&" and month="&p_month&" order by  day asc"
    'sqlbb="SELECT * from kcbb where class="&request("zclass")
	set rsbb=server.createobject("adodb.recordset")
    rsbb.open sqlbb,conne,1,1
    if rsbb.eof and rsbb.bof then 
	  response.write "此月报表未生成"
	else
			response.write " <TABLE  width=""100%""><tr>"
			 response.write " <td><div align=center>设备管理系统</div></td>"
			response.write " </tr></TABLE>"
			response.write " <TABLE  width=""100%""><tr>"
			 response.write " <td><div align=center>"&rsbb("year")&"年"&rsbb("month")&"月份员工教育培训报表</div></td>"
			response.write " </tr></TABLE>"
			response.write " <TABLE  width=""100%""><tr>"
			 response.write " <td>&nbsp;&nbsp;</TD><TD><div align=left>单位：</div></td>"
			 response.write " <td><div align=right>"&rsbb("year")&"年"&rsbb("month")&"月</div></td><td>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			response.write " </tr></TABLE>"
        	response.write "<table border=1 cellpadding=0 cellspacing=0 width=""100%"">"
			response.write " <tr>"
			response.write "  <td><div align=center>时间</div></td>"
			response.write "  <td ><div align=center>培训内容摘要</div></td>"
			response.write "  <td ><div align=center>培训对象</div></td>"
			response.write "  <td ><div align=center>计划人数</div></td>"
			response.write "  <td ><div align=center>实际人数</div></td>"
			response.write "  <td ><div align=center>计划课时</div></td>"
			response.write "  <td ><div align=center>实际课时</div></td>"
			response.write "  <td ><div align=center>培训率</div></td>"
			response.write "  <td ><div align=center>合格率</div></td>"
			response.write "  <td ><div align=center>授课人</div></td>"
			response.write "  <td><div align=center>备注</div></td>"
			response.write " </tr>"
       do while not rsbb.eof
			response.write " <tr >"
			response.write "  <td><div align=center>"&rsbb("month")&"."&rsbb("day")&"</div></td>"
			response.write "  <td>"&rsbb("body")&"</td>"
			response.write "  <td>"&rsbb("pxdx")&"</td>"
			response.write "  <td>"&rsbb("numb")&"</td>"
			response.write "  <td>"&rsbb("sjnumb")&"</td>"
			response.write "  <td><div align=center>"&rsbb("ks")&"</div></td>"
			response.write "  <td><div align=center>"&rsbb("sjks")&"</div></td>"
			response.write "  <td><div align=center>"&rsbb("pxl")&"</div></td>"
			response.write "  <td><div align=center>"&rsbb("hgl")&"</div></td>"
			response.write "  <td><div align=center>"&rsbb("skrname")&"</div></td>"
			response.write "  <td>"&ssbzh(rsBB("SSBZ"))&rsbb("bz")&"&nbsp;</td>"
			response.write " </tr>"
		 rsbb.movenext
		 loop
			response.write "</table>"
	   end if
	rsbb.close
	set rsbb=nothing
end function




function jqnumb(zfc)
'返回字符串的数字
	dim lenzfc,i,numb
	'dim leftynumb,leftyzfc
	lenzfc=len(zfc)
	for i=1 to lenzfc
		
		if IsNumeric(mid(zfc,i,1)) then 
			numb=cint(numb&mid(zfc,i,1))
			
		end if
	next
	jqnumb=numb
	response.write jqnumb
'response.write leftynumb&"<br>"&leftyzfc
end function

function jqtext(zfc)
'返回字符串的文字
	dim lenzfc,i,text
	'dim leftynumb,leftyzfc
	lenzfc=len(zfc)
	for i=1 to lenzfc
		
		if IsNumeric(mid(zfc,i,1)) then 
		else
			text=text&mid(zfc,i,1)
		end if
	next
	jqtext=replace(text,"人","")
	response.write jqtext
'response.write leftynumb&"<br>"&leftyzfc
end function

Call CloseConn
%>