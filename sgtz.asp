<%@language=vbscript codepage=936 %>
<%
'Option Explicit
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
dim title,record,pgsz,total,page,start,rowcount,url,ii,zxzz
dim id,scontent,rsdel,sqldel
url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>事故台账</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.wh.value==''){" & vbCrLf
dwt.out "      alert('位号不能为空！');" & vbCrLf
dwt.out "   document.form1.wh.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
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

sub add()
dim rscj,sqlcj
 	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加事故记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='sgtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	if session("level")=0 then 
		dwt.out"<select name='sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
		sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
		dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
		rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		
		dwt.out"</select>"  	 
	else 	 
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
		dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
	end if 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故位号名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wh' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='address' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 115px' >事故名称及内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
'	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wh' type='text' value='"&rsedit("wh")&"'>"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>事故类别:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<select name='sgclass' style='WIDTH: 175px' size='1'>"
	dwt.out"<option value='1'>设备事故</option>"
	dwt.out"<option value='2'>操作事故</option>"
	dwt.out"<option value='3'>责任事故</option>"
    dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>事故时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='createdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)'  >"
		'dwt.out "<br/>日期时间格式为:2008-02-02 08:22:22"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>责任人:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ren' type='text' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>责任人处理:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='cljg' type='text' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故发生详细经过及主要原因:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jg></TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>防范措施:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=clcs></TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>备注:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=bz></TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf



	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  
end sub	

sub saveadd()    
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sgtz" 
	rsadd.open sqladd,connb,1,3
	rsadd.addnew
	on error resume next
	rsadd("sscj")=Trim(Request("sscj"))
	rsadd("wh")=request("wh")
	rsadd("class")=Trim(request("sgclass"))
	rsadd("address")=request("address")
	rsadd("createdate")=request("createdate")
	rsadd("ren")=request("ren")
	rsadd("jg")=request("jg")
	rsadd("clcs")=request("clcs")
	rsadd("cljg")=request("cljg")
	rsadd("bz")=request("bz")
	
	rsadd.update
	rsadd.close
	set rsadd=nothing
	dwt.out"<Script Language=Javascript>location.href='sgtz.asp';</Script>"
end sub


sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sgtz where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connb,1,3
	on error resume next
	'rsedit("sscj")=Trim(Request("_sscj"))
	rsedit("wh")=request("wh")
	rsedit("address")=request("address")
	rsedit("class")=Trim(request("sgclass"))
	rsedit("createdate")=request("createdate")
	'rsedit("completedate")=request("completedate")
	rsedit("ren")=request("ren")
	rsedit("cljg")=request("cljg")
	rsedit("jg")=request("jg")
	'rsedit("yy")=request("yy")
	rsedit("clcs")=request("clcs")
	rsedit("bz")=request("bz")
	'	rsedit("cjyj")=request("cjyj")
	'rsedit("fcyj")=request("fcyj")
	'rsedit("shyj")=request("shyj")
	'rsedit("cldyj")=request("cldyj")

	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
	id=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from sgtz where id="&id
	rsdel.open sqldel,connb,1,3
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub


sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sgtz where id="&id
	rsedit.open sqledit,connb,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 370px;padding-top:50px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑事故台账记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='sgtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>属所车间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px'  value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' >"& vbCrLf
	dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' type='hidden' name='sscj' value='"&sscjh(rsedit("sscj"))&"'  >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故位号名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wh' type='text' value='"&rsedit("wh")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='address' type='text' value='"&rsedit("address")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 115px' >事故名称及内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
'	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wh' type='text' value='"&rsedit("wh")&"'>"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>事故类别:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<select name='sgclass' style='WIDTH: 175px' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("class")=1 then dwt.out "selected"
	dwt.out ">设备事故</option>"
	dwt.out"<option value='2'"
	if rsedit("class")=2 then dwt.out "selected"
	dwt.out ">操作事故</option>"
	dwt.out"<option value='3'"
	if rsedit("class")=3 then dwt.out "selected"
	dwt.out ">责任事故</option>"
    dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>事故时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='createdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' value='"&rsedit("createdate")&"'>"
	dwt.out "<br/>日期时间格式为:2008-02-02 08:22:22"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>责任人:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ren' type='text' value='"&rsedit("ren")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>责任人处理:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='cljg' type='text' value='"&rsedit("ren")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>事故发生经过及主要原因:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jg>"&rsedit("jg")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>防范措施:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=clcs>"&rsedit("clcs")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>备注:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=bz>"&rsedit("bz")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 完 成 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  
	rsedit.close
	set rsedit=nothing
end sub


sub main()
	'sql="SELECT * from zg ORDER BY id DESC"
	sql="SELECT * from sgtz"
	if keys<>"" then 
		sql=sql&" where wh like '%"&keys&"%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
		
        sql=sql&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	sql=sql&" ORDER BY sscj aSC,createdate desc"

	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>事故台账"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

	for sscji=1 to 5 
		sqlsqltotal="SELECT count(id) FROM sgtz WHERE sscj="&sscji
		numb=numb&sscjh_d(sscji)&"<span style='color:#006600;'>"&connb.Execute(sqlsqltotal)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
	next
	
	sqltotal="SELECT count(id) FROM sgtz "
	totall= "<span style='color:#006600;'>"&connb.Execute(sqlsqltotal)(0)&"</span>" 
	dwt.out "<div class='pre'>本月:"&numb&"合计:"&totall&"</div>"& vbCrLf
	call search()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connb,1,1
	if rs.eof and rs.bof then 
	   message "未找到相关内容"
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>事故位号名称</div></td>"
		'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>事故名称</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>事故类型</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>发生日期</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>责任人</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
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
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&xh_id&"</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"
			dwt.out sscjh_d(rs("sscj"))&"</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><a href=sgtz_view.asp?id="&rs("id")&"  target=_blank>"
			if keys<>"" then 
			      dwt.out searchH(uCase(rs("wh")),keys)
			else
   			      dwt.out rs("wh")
			end if 	  
	  
			dwt.out "</a></td>"
			dim sgclass
			if rs("class")=1 then sgclass="设备事故"
			if rs("class")=2 then sgclass="操作事故"
			if rs("class")=3 then sgclass="责任事故"
			'dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rs("name")&"</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&sgclass&"</div></td>"
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rs("createdate")&"</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rs("ren")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"
			call editdel(rs("id"),rs("sscj"),"sgtz.asp?action=edit&id=","sgtz.asp?action=del&id=")
			dwt.out "</div></td></tr>"
			 RowCount=RowCount-1
          rs.movenext
		loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" or sscjid<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 
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
	dwt.out " <form method='Get' name='SearchForm' action='sgtz.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "  <a href=""sgtz.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
	dwt.out "<strong>位号搜索：</strong>" & vbCrLf
	dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	dwt.out "<font color='0066CC'> 查看所属车间的相关内容：</font>"
	dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='sgtz.asp?sscj="&rscj("levelid")&"'"
	if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
			dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext	
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf
	dwt.out "</div></div></form>" & vbCrLf
end sub





Call CloseConn
%>