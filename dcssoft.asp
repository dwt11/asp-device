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
dim url,record,pgsz,total,page,start,rowcount,ii,pagename
'urljx="dcssoft.asp?action=dcsjx"
'urlgh="dcssoft.asp?action=dcsgh"
dim keys,sscjid
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
url="dcssoft.asp"

'if request("action")="dcsgh" or request("action")="addgh" or request("action")="editgh" then pagename="DCS更换记录"
'if request("action")="dcsjx"  or request("action")="addjx" or request("action")="editjx" then pagename="DCS检修记录"
'if request("action")="dcssoft" or request("action")="addsoft" or request("action")="editsoft"  then pagename="DCS软件工作记录"
'if request("action")="" then pagename="DCS更换记录"


dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统DCS\PLC更换检修记录管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkaddgh(){" & vbCrLf
dwt.out " if(document.form1.dcsgh_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.dcsgh_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.dcsgh_sbname.value==''){" & vbCrLf
dwt.out "      alert('设备名称不能为空！');" & vbCrLf
dwt.out "   document.form1.dcsgh_sbname.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "function checkaddjx(){" & vbCrLf
dwt.out " if(document.form1.dcsjx_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.dcsjx_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.dcsjx_body.value==''){" & vbCrLf
dwt.out "      alert('检修内容不能为空！');" & vbCrLf
dwt.out "   document.form1.dcsjx_body.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "function checkaddsoft(){" & vbCrLf
dwt.out " if(document.form1.dcssoft_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.dcssoft_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.dcssoft_body.value==''){" & vbCrLf
dwt.out "      alert('作业内容不能为空！');" & vbCrLf
dwt.out "   document.form1.dcssoft_body.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "function checksearch(){" & vbCrLf
dwt.out " if(document.searchform.search_class.value==''){" & vbCrLf
dwt.out "      alert('请选择搜索类型！');" & vbCrLf
dwt.out "   document.searchform.search_class.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf

dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
    if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	
'if request("action")="dcsgh" then call dcsgh
'if request("action")="" then call dcsgh
'if request("action")="dcsjx" then dcsjx
'if request("action")="deljx"  then deljx
'if request("action")="addjx" then addjx
'if request("action")="saveaddjx" then saveaddjx
'if request("action")="editjx" then editjx
'if request("action")="saveeditjx" then saveeditjx
'if request("action")="dcssoft" then dcssoft
'if request("action")="addsoft" then addsoft
'if request("action")="editsoft" then editsoft
'if request("action")="saveaddsoft" then saveaddsoft
'if request("action")="saveeditsoft" then saveeditsoft
'if request("action")="delsoft" then delsoft


'sub addgh()
'dim sqlcj,rscj
'   dwt.out"<br><br><br><form method='post' action='dcssoft.asp' name='form1' onsubmit='javascript:return checkaddgh();'>"
'   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
'   dwt.out"<tr class='title'><td height='22' colspan='2'>"
'   dwt.out"<div align='center'><strong>添加DCS更换记录</strong></div></td>    </tr>"
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'>"
'  if session("level")=0 then 
'	dwt.out"<select name='dcsgh_sscj' size='1'>"
'    dwt.out"<option >选择所属车间</option>"
'    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'    dwt.out"</select></td></tr>  "  	 
'  else 	 
'     dwt.out"<input name='dcsgh_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
'      dwt.out"<input name='dcsgh_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
'
' end if 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
'	 dwt.out"<strong>设备名称：</strong></td>"
'	 dwt.out"<td width='88%' class='tdbg'><input name='dcsgh_sbname' type='text'></td>    </tr>   "
'	 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>规格型号：</strong></td> "
'	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsgh_ggxh' ></td></tr> "
'	 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>安装位置：</strong></td> "
'	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsgh_azwz'></td></tr> "
'	 
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>更换原因：</strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsgh_ghyy'></td></tr>  "   
'   
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>损坏时间：</strong></td>"      
'   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'><script>"
'   dwt.out"<input name='dcsgh_shdate' type='text' value="&now()&" >"
'   dwt.out"<a href='#' onClick=""popUpCalendar(this, dcsgh_shdate, 'yyyy-mm-dd'); return false;"">"
'   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
'	
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>更换时间：</strong></td>"      
'   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'><script>"
'   dwt.out"<input name='dcsgh_ghdate' type='text' value="&now()&" >"
'   dwt.out"<a href='#' onClick=""popUpCalendar(this, dcsgh_ghdate, 'yyyy-mm-dd'); return false;"">"
'   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
'
'   	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>更换人：</strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsgh_ghrname'></td></tr>  "   
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsgh_bz'></td></tr>  "   
'
'	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
'	dwt.out"<input name='action' type='hidden' id='action' value='saveaddgh'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
'	dwt.out"</table></form>"
'end sub	
'
'sub saveaddgh()  
'dim sqladd,rsadd  
'	 on error resume next '保存
'      set rsadd=server.createobject("adodb.recordset")
'      sqladd="select * from dcsgh" 
'      rsadd.open sqladd,conndcs,1,3
'      rsadd.addnew
'      rsadd("sscj")=Trim(Request("dcsgh_sscj"))
'      rsadd("sbname")=request("dcsgh_sbname")
'      rsadd("ggxh")=Trim(request("dcsgh_ggxh"))
'      rsadd("azwz")=request("dcsgh_azwz")
'      rsadd("ghyy")=request("dcsgh_ghyy")
'      rsadd("shdate")=request("dcsgh_shdate")
'      rsadd("ghdate")=request("dcsgh_ghdate")
'      rsadd("ghrname")=request("dcsgh_ghrname")
'      rsadd("bz")=request("dcsgh_bz")
'      rsadd.update
'      rsadd.close
'      set rsadd=nothing
'	  dwt.out"<Script Language=Javascript>location.href='dcssoft.asp?action=dcsgh'<Script>"
'end sub
'
'
'sub addjx()
'dim sqlcj,rscj
'   dwt.out"<br><br><br><form method='post' action='dcssoft.asp' name='form1' onsubmit='javascript:return checkaddjx();'>"
'   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
'   dwt.out"<tr class='title'><td height='22' colspan='2'>"
'   dwt.out"<div align='center'><strong>添加DCS检修记录</strong></div></td>    </tr>"
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'>"
'  if session("level")=0 then 
'	dwt.out"<select name='dcsjx_sscj' size='1'>"
'    dwt.out"<option >请选择所属车间</option>"
'    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'    dwt.out"</select></td></tr>  "  	 
'  else 	 
'     dwt.out"<input name='dcsjx_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
'      dwt.out"<input name='dcsjx_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
'
' end if 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
'	 dwt.out"<strong>检修原因：</strong></td>"
'	 dwt.out"<td width='88%' class='tdbg'><input name='dcsjx_jxyy' type='text'></td>    </tr>   "
'	 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>检修内容：</strong></td> "
'	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsjx_body' ></td></tr> "
'	 
'	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>检修人：</strong></td> "
'	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsjx_jxrname'></td></tr> "
'	    
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>检修时间：</strong></td>"      
'   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'><script>"
'   dwt.out"<input name='dcsjx_jxdate' type='text' value="&now()&" >"
'   dwt.out"<a href='#' onClick=""popUpCalendar(this, dcsjx_jxdate, 'yyyy-mm-dd'); return false;"">"
'   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
'	
'	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
'    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcsjx_bz'></td></tr>  "   
'
'	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
'	dwt.out"<input name='action' type='hidden' id='action' value='saveaddjx'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
'	dwt.out"</table></form>"
'end sub	

sub add()
   dim sqlcj,rscj
   dwt.out"<br><br><br><form method='post' action='dcssoft.asp' name='form1' onsubmit='javascript:return checkaddsoft();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>添加DCS检修记录</strong></div></td>    </tr>"
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='dcssoft_sscj' size='1'>"
    dwt.out"<option >请选择所属车间</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
     dwt.out"<input name='dcssoft_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
      dwt.out"<input name='dcssoft_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf

 end if 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>作业原因：</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><textarea name='dcssoft_zyyy' cols='20' rows='5'></textarea></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业内容：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><textarea name='dcssoft_body' cols='20' rows='5'></textarea></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业人：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcssoft_zyrname'></td></tr> "
	    
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业时间：</strong></td>"      
   dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='dcssoft_zydate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcssoft_bz'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	



sub saveadd()  
dim sqladd,rsadd  
	 on error resume next '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from dcssoft" 
      rsadd.open sqladd,conndcs,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("dcssoft_sscj"))
      rsadd("zyyy")=request("dcssoft_zyyy")
      rsadd("body")=Trim(request("dcssoft_body"))
      rsadd("zyrname")=request("dcssoft_zyrname")
      rsadd("zydate")=request("dcssoft_zydate")
      rsadd("bz")=request("dcssoft_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='dcssoft.asp';</Script>"
end sub



sub saveedit()  
dim rsedit,sqledit  
	 on error resume next '保存
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from dcssoft where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conndcs,1,3
      rsedit("sscj")=Trim(Request("dcssoft_sscj"))
      rsedit("zyyy")=request("dcssoft_zyyy")
      rsedit("body")=Trim(request("dcssoft_body"))
      rsedit("zyrname")=request("dcssoft_zyrname")
      rsedit("zydate")=request("dcssoft_zydate")
      rsedit("bz")=request("dcssoft_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub del()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from dcssoft where id="&id
  rsdel.open sqldel,conndcs,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub


sub edit()
   dim sqledit,rsedit,id
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from dcssoft where id="&id
   rsedit.open sqledit,conndcs,1,1
   dwt.out"<br><br><br><form method='post' action='dcssoft.asp' name='form1' onsubmit='javascript:return checkaddsoft();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>编辑DCS软件操作记录</strong></div></td>    </tr>"
     
     dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'><input name='dcssoft_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out"<input name='dcssoft_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>作业原因：</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><textarea name='dcssoft_zyyy' cols='20' rows='5'>"&rsedit("zyyy")&"</textarea></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业内容：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><textarea name='dcssoft_body' cols='20' rows='5'>"&rsedit("body")&"</textarea></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业人：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcssoft_zyrname' value="&rsedit("zyrname")&"></td></tr> "
   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作业时间：</strong></td>"      
   dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='dcssoft_zydate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("zydate")&"'>"
   dwt.out"</td></tr>"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='dcssoft_bz' value="&rsedit("bz")&"></td></tr>  "   
	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

end sub






sub MAIN()
	dim title,sql,rs
	sql="SELECT * from dcssoft"
	if keys<>"" then 
		sql=sql&" where zyyy like '%" &keys& "%' or body like '%" &keys& "%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
		sql=sql&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	sql=sql&"  ORDER BY zydate DESC"

	
	
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>DCS软件工作记录 "&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	call search()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conndcs,1,1
	if rs.eof and rs.bof then 
		dwt.out "<p align='center'>未添加DCS软件工作记录</p>" & vbCrLf
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr  class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>作业原因</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>作业内容</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>作业人</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>作业时间</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>备注</div></td>"& vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"& vbCrLf
		dwt.out "    </tr>"
		record=rs.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
		end if 
		rs.PageSize = Cint(PgSz) 
		total=int(record/PgSz*-1)*-1
		page=Request("page")
		if page="" Then
		  page = 1
		else
		 page=page+1
		 page=page-1
		end if
		if page<1 Then 
		  page=1
		end if
		rs.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rs.PageSize
		do while not rs.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			'dwt.out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&xh&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&sscjh_d(rs("sscj"))&"</td>"& vbCrLf
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&searchH(uCase(rs("zyyy")),keys)&"</td>"& vbCrLf
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&searchH(uCase(rs("body")),keys)&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("zyrname")&"&nbsp;</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rs("zydate")&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rs("bz")&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"
			call editdel(rs("id"),rs("sscj"),"dcssoft.asp?action=edit&id=","dcssoft.asp?action=del&id=")
			dwt.out "</div></td></tr>"& vbCrLf
			RowCount=RowCount-1
		rs.movenext
		loop
		dwt.out "</table>"& vbCrLf
		call showpage1(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end sub


dwt.out "</body></html>"

sub search()
	
	
	
	dim sqlcj,rscj
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='dcssoft.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "<a href=""dcssoft.asp?action=add"">添加工作记录</a>"
	dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if keys<>"" then 
		dwt.out "value='"&keys&"'"
		dwt.out "/>" & vbCrLf
	else
		dwt.out "value='输入搜索的内容'"
		dwt.out " onblur=""if(this.value==''){this.value='输入搜索的内容'}"" onfocus=""this.value=''""/>" & vbCrLf
	end if    
	dwt.out "  <input type='Submit' name='Submit'  value='搜索'/>" & vbCrLf
	dwt.out "  <input type='hidden' name='search' value='keys'/>" & vbCrLf
	
	dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
	set rscj=server.createobject("adodb.recordset")
	rscj.open sqlcj,conn,1,1
	do while not rscj.eof
		dwt.out"<option value='dcssoft.asp?sscj="&rscj("levelid")&"'"
		if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
	
		dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	dwt.out "     </select>	" & vbCrLf
	
dwt.out "</div></div>"

end sub
Call CloseConn
%>