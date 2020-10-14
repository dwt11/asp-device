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
dim sqlqxtb,rsqxtb,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
dim sqlhuiyiluoshi_fk,rshuiyiluoshi_fk,huiyiluoshi_id
if Request("action")="add" then  call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then	call del
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统缺陷整改通知管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf

dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf


sub add()
   dwt.out"<br><br><br><form method='post' action='anquangs_fk.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>添加反馈</strong></div></td>    </tr>"
        dwt.out"<tr><td>属所车间:</td>"& vbCrLf
	dwt.out"<td>"
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
		dwt.out"</td></tr>"& vbCrLf
	


	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>时&nbsp;&nbsp;间：</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='huiyiluoshi_fk_date' type='text' value="&now()&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this,huiyiluoshi_fk_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
	 dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=huiyiluoshi_fk_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
      dwt.out"</iframe>  <input type='hidden' name='huiyiluoshi_fk_body' value=''>"
    dwt.out"</td></tr>  "   
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'> <input name='huiyiluoshi_id' type='hidden' id='huiyiluoshi_id' value='"&request("huiyiluoshi_id")&"'>   <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='anquangs_luoshi.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from anquangs_fk" 
      rsadd.open sqladd,connaq,1,3
      rsadd.addnew
            rsadd("huiyiluoshi_id")=request("huiyiluoshi_id")
rsadd("huiyiluoshi_fk_sscj")=session("levelclass")
      rsadd("huiyiluoshi_fk_body")=Trim(request("huiyiluoshi_fk_body"))
      rsadd("huiyiluoshi_fk_date")=request("huiyiluoshi_fk_date")
      
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='anquangs_luoshi.asp';</Script>"
end sub

sub edit()
   huiyiluoshi_id=ReplaceBadChar(Trim(request("huiyiluoshi_id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from anquangs_fk where huiyiluoshi_id="&huiyiluoshi_id&" and huiyiluoshi_fk_sscj="&session("levelclass")
   rsedit.open sqledit,connaq,1,1

   dwt.out"<form method='post' action='anquangs_fk.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>编 辑 反 馈</strong></div></td>    </tr>"
   dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>时&nbsp;&nbsp;间：</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='huiyiluoshi_fk_date' type='text' value="&rsedit("huiyiluoshi_fk_date")&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this,huiyiluoshi_fk_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	  
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
	scontent=rsedit("huiyiluoshi_fk_body")
	 dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=huiyiluoshi_fk_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      dwt.out"</iframe><textarea name='huiyiluoshi_fk_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    dwt.out"</td></tr>  "   

	 
    dwt.out"<tr> <td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input name='huiyiluoshi_id' type='hidden' id='huiyiluoshi_id' value='"&request("huiyiluoshi_id")&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='anquangs_luoshi.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from anquangs_fk where huiyiluoshi_id="&ReplaceBadChar(Trim(request("huiyiluoshi_id")))&" and huiyiluoshi_fk_sscj="&session("levelclass")

rsedit.open sqledit,connaq,1,3
rsedit("huiyiluoshi_fk_body")=Trim(request("huiyiluoshi_fk_body"))
rsedit("huiyiluoshi_fk_date")=request("huiyiluoshi_fk_date")
rsedit.update
rsedit.close
	  dwt.out"<Script Language=Javascript>location.href='anquangs_luoshi.asp';</Script>"
	
end sub




sub del()
huiyiluoshi_id=request("huiyiluoshi_id")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from anquangs_fk where huiyiluoshi_id="&huiyiluoshi_id&" and huiyiluoshi_fk_sscj="&session("level")
rsdel.open sqldel,connaq,1,3
	  dwt.out"<Script Language=Javascript>location.href='anquangs_luoshi.asp';</Script>"
'rsdel.close
set rsdel=nothing  

end sub




dwt.out "</body></html>"



Call CloseConn
%>