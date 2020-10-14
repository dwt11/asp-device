<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->


<%
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统月计划内容显示</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checkadd(){" & vbCrLf
response.write " if(document.form1.yjh_sscj.value==''){" & vbCrLf
response.write "      alert('请选择所属支部！');" & vbCrLf
response.write "   document.form1.yjh_sscj.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf

response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf

response.write "   <td height='22' colspan='2' align='center'><strong>党支部月计划</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""d_yjhzj.asp"">月计划管理首页</a>&nbsp;|&nbsp;<a href=""d_yjh_view.asp?action=addyjh"">添加月计划</a>&nbsp;|&nbsp;<a href=""d_yzj_view.asp?action=addyzj"">添加月总结</a></td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

if request("action")="yjh" then call yjh()
if request("action")="addyjh" then call addyjh()
if request("action")="saveaddyjh" then call saveaddyjh()
if request("action")="del" then call del()
if request("action")="edit" then call edit()
if request("action")="saveedit" then saveedit()
sub addyjh()
dim ii
dim rscj,sqlcj,rsbz,sqlbz,sql,rs
   response.write"<form method='post' action='d_yjh_view.asp' name='form1'  onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>添加月计划</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属支部： </strong></td>"& vbCrLf      
    response.write"<td width='88%' class='tdbg'>"& vbCrLf
  'if session("level")=0 then 
	response.write"<select name='yjh_sscj' size='1'>"
    response.write"<option >请选择所属支部</option>"
		response.write"<option  value='1'>维修一党织部</option>"
		response.write"<option  value='2'>维修二党织部</option>"
		response.write"<option  value='3'>机关党织部</option>"
		response.write"<option  value=4>维修三党支部</option>"
		response.write"<option  value=5>维修四党支部</option>"



'    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
    response.write"</select></td></tr>  "  	 
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>月计划日期：</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='yjh_date' type='text' value="&year(now())&"-"&month(now())&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, yjh_date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>月计划内容：</strong></td>"
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=yjh_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe>  <input type='hidden' name='yjh_body' value=''>"
	  response.write"</td></tr>   "
	 
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveaddyjh'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveaddyjh()    
	  dim year1,month1,day1'保存\
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from d_yjh" 
      rsadd.open sqladd,conna,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("yjh_sscj"))
      year1=year(Trim(Request("yjh_date")))
	  month1=month(Trim(Request("yjh_date")))
	  if len(month1)<>2 then month1="0"&month1
      rsadd("month")=month1
	  rsadd("year")=year1
  	  dim yjh_body
     yjh_body=request("yjh_body")
	  if yjh_body="" then   yjh_body=" "
     rsadd("body")=yjh_body
      rsadd("userid")=session("userid")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='d_yjhzj.asp';</Script>"
end sub



sub saveedit()    
	  dim year1,month1,day1'保存\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from d_yjh where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conna,1,3
      rsedit("sscj")=Trim(Request("yjh_sscj"))
      year1=year(Trim(Request("yjh_date")))
	  month1=month(Trim(Request("yjh_date")))
	  if len(month1)<>2 then month1="0"&month1
      rsedit("month")=month1
	  rsedit("year")=year1
      rsedit("body")=Trim(request("yjh_body"))

      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub edit()
    dim scontent
   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from d_yjh where id="&id
   rsedit.open sqledit,conna,1,1

   response.write"<form method='post' action='d_yjh_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>编辑月计划</strong></div></td>    </tr>"
	
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属支部： </strong></td>"   & vbCrLf   
     
	  		dim sszb
		if rsedit("sscj")=1 then sszb="维修一党织部"
		if rsedit("sscj")=2 then sszb="维修二党织部"
		if rsedit("sscj")=3 then sszb="机关党织部"
		if rsedit("sscj")=4 then sszb="维修三党支部"
		if rsedit("sscj")=5 then sszb="维修四党支部"

	 response.write"<td width='88%' class='tdbg'><input name='yzj_sscj'  disabled='disabled'  type='text' value='"&sszb&"'></td></tr>"& vbCrLf
     response.write"<input name='yjh_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf
	    
		response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>月计划日期：</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='yjh_date' type='text' value="&rsedit("year")&"-"&rsedit("month")&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, yjh_date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	scontent=rsedit("body")
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=yjh_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe><textarea name='yjh_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    response.write"</td></tr>  "   

	
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub

sub del()
 dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from d_yjh where id="&request("id")
  rsdel.open sqldel,conna,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub yjh()
   dim rsyjh,sqlyjh,rs,sql
   '显示车间级的月计划
      sqlyjh="SELECT * from d_yjh where sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rsyjh=server.createobject("adodb.recordset")
      rsyjh.open sqlyjh,conna,1,1
    	  	dim sszb
		if request("sscj")=1 then sszb="维修一党织部"
		if request("sscj")=2 then sszb="维修二党织部"
		if request("sscj")=3 then sszb="机关党织部"
        if request("sscj")=4 then sszb="维修三党支部"
		if request("sscj")=5 then sszb="维修四党支部"

         response.write "<br><table width=""90%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			 
             response.write " <tr class=""title""><td height=40><div align=center><strong>"&sszb&request("year")&"年"&request("month")&"月份工作计划</strong></div>"
             response.write "</td></tr>"
             response.write "<tr class=""tdbg"">"
			 response.write "<td><table width=""90%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"
			 response.write rsyjh("body")
			 response.write "</td></tr></table></td>"
             response.write "</tr></table><br>"		
  rsyjh.close
  set rsyjh=nothing
end sub


response.write "</body></html>"
Call CloseConn
%>