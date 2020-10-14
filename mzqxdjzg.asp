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
dim sqlqxdj,rsqxdj,title,record,pgsz,total,page,start,rowcount,url,ii,zxzz
dim rsadd,sqladd,qxdjid,rsedit,sqledit,scontent,rsdel,sqldel,tyzk,id
url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统每周缺陷记录管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.qxdj_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.qxdj_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.qxdj_wh.value==''){" & vbCrLf
dwt.out "      alert('位号不能为空！');" & vbCrLf
dwt.out "   document.form1.qxdj_wh.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")


select case action
  case "add"
       'if truepagelevelh(session("groupid"),1,session("pagelevelid")) then 
	   call add
  case "saveadd"
    call saveadd
  case "edit"
	'if truepagelevelh(session("groupid"),2,session("pagelevelid")) then
	 call edit
  case "saveedit"
    call saveedit
  case "fk"
       'if truepagelevelh(session("groupid"),1,session("pagelevelid")) then 
	   call fk
  case "savefk"
    call savefk
  case "editfk"
	'if truepagelevelh(session("groupid"),2,session("pagelevelid")) then
	 call editfk
  case "saveeditfk"
    call saveeditfk
  case "del"
    'if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
	call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then 
	call main
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
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加每周缺陷记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='mzqxdjzg.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	'if session("level")=0 then 
		dwt.out"<select name='qxdj_sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
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
'	else 	 
'		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
'		dwt.out"<input name='qxdj_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
'	end if 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >位   号:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_wh>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >缺陷内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_body >请添写缺陷内容</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	if isnull(session("jcdate1")) or session("jcdate1")="" then
	 jcdate1=date()-Weekday(Date())+2
    else
	 jcdate1=session("jcdate1")
    end if 
	if isnull(session("jcdate2"))  or session("jcdate2")="" then
	 jcdate2=date()+8-Weekday(Date())
    else
	 jcdate2=session("jcdate2")
    end if 
	
'dwt.out 	session("jcdate2")&"---------"
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检查周:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jcdate1' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&jcdate1&"'><br> <input name='jcdate2' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&jcdate2&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>整改结果:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'    dwt.out"<select name='qxdj_zgjg' style='WIDTH: 175px' size='1'>"
'	dwt.out"<option value='true'>已整改</option>"
'    dwt.out"<option value='false'>未整改</option>"
'    dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >缺陷内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_zgbody ></TEXTAREA>"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>督办人:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_dbname >"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
'
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>整改时间:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'    dwt.out"<input name='qxdj_zgdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value=''>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
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
	sqladd="select * from mzqxdjzg" 
	rsadd.open sqladd,connb,1,3
	rsadd.addnew
	on error resume next
	rsadd("sscj")=Trim(Request("qxdj_sscj"))
	rsadd("wh")=request("qxdj_wh")
	rsadd("body")=Trim(request("qxdj_body"))
	rsadd("jcdate1")=request("jcdate1")
	rsadd("jcdate2")=request("jcdate2")
	session("jcdate1")=request("jcdate1")
	session("jcdate2")=request("jcdate2")
	'rsadd("update")=now()
	rsadd.update
	rsadd.close
	set rsadd=nothing
	dwt.out"<Script Language=Javascript>location.href='mzqxdjzg.asp';</Script>"
end sub


sub del()
	qxdjid=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from mzqxdjzg where id="&qxdjid
	rsdel.open sqldel,connb,1,3
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub

sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from mzqxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connb,1,3
	'on error resume next
	rsedit("sscj")=Trim(Request("qxdj_sscj"))
	rsedit("wh")=request("qxdj_wh")
	rsedit("body")=Trim(request("qxdj_body"))
	rsedit("jcdate1")=request("jcdate1")
	rsedit("jcdate2")=request("jcdate2")
	rsedit("qrname")=request("qrname")
	rsedit("qrdate")=request("qrdate")

	'rsedit("update")=now()
	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from mzqxdjzg where id="&id
	rsedit.open sqledit,connb,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑每周缺陷记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='mzqxdjzg.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
		dwt.out"<select name='qxdj_sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
		'dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
		sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
		dwt.out"<option value='"&rscj("levelid")&"' "
		if rscj("levelid")=rsedit("wh") then dwt.out "selected"
		dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
		rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >位号:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='qxdj_wh' type='text' value='"&rsedit("wh")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >缺陷内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_body >"&rsedit("body")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检查周:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jcdate1' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("jcdate1")&"'><br><input name='jcdate2' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("jcdate2")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>整改结果:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'    dwt.out"<select name='qxdj_zgjg' style='WIDTH: 175px' size='1'>"
'	dwt.out"<option value='true'"
'	if rsedit("zgjg") then dwt.out "selected"
'	dwt.out ">已整改</option>"
'    dwt.out"<option value='false'"
'	if rsedit("zgjg")=false then dwt.out "selected"
'	dwt.out ">未整改</option>"
'    dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
'
'
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >缺陷内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_zgbody >"&rsedit("zgbody")&"</TEXTAREA>"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
'
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>督办人:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_dbname value='"&rsedit("dbname")&"'>"& vbCrLf
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>整改时间:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
'    dwt.out"<input name='qxdj_zgdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("zgdate")&"'>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
sub savefk()    
	if request("qxdj_zgjg")=10 then 
		set rsadd=server.createobject("adodb.recordset")
		sqladd="select * from qxdjzg" 
		rsadd.open sqladd,connb,1,3
		rsadd.addnew
		
		set rsedit=server.createobject("adodb.recordset")
		sqledit="select * from mzqxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
		rsedit.open sqledit,connb,1,1
	
		rsadd("sscj")=rsedit("sscj")
		rsadd("wh")=rsedit("wh")
		rsadd("body")=rsedit("body")
		rsadd("cxdate")=rsedit("jcdate1")
		rsadd("update")=now()
		tzid=rsadd("ID")
		rsedit.close
		set rsedit=nothing
		rsadd.update
		rsadd.close
		set rsadd=nothing
	
	
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from mzqxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
	rsdel.open sqldel,connb,1,3
	set rsdel=nothing  

		dwt.out"<Script Language=Javascript>window.location.href='qxdjzg.asp?action=edit&id="&tzid&"';</Script>"
	
	else
		set rsedit=server.createobject("adodb.recordset")
		sqledit="select * from mzqxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
		rsedit.open sqledit,connb,1,3
		'on error resume next
		rsedit("zgjg")=Trim(Request("qxdj_zgjg"))
		rsedit("zgzt")=Trim(Request("qxdj_zgjg"))
		rsedit("zgbody")=request("qxdj_zgbody")
		rsedit("zgname")=Trim(request("qxdj_zgname"))
		rsedit("zgdate")=request("qxdj_zgdate")
		rsedit("qrname")=Trim(request("qxdj_qrname"))
		rsedit("qrdate")=request("qxdj_qrdate")
		rsedit.update
		rsedit.close
		set rsedit=nothing
		dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
	end if
end sub

sub fk()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from mzqxdjzg where id="&id
	rsedit.open sqledit,connb,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加反馈</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='mzqxdjzg.asp' name='form1'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf

	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>整改状态:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<select name='qxdj_zgjg' style='WIDTH: 175px' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("zgjg") then dwt.out "selected"
	dwt.out ">已整改</option>"
    dwt.out"<option value='0'"
	if rsedit("zgjg")=false then dwt.out "selected"
	dwt.out ">未整改</option>"
    dwt.out"<option value='10'>不具备整改条件</option>"
    dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >整改内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_zgbody >"&rsedit("zgbody")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
    if isnull(rsedit("zgname")) or rsedit("zgname")="" then 
	  zgname=session("UserName1")
	else
	  zgname=rsedit("zgname")  
	end if 
  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>整改人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_zgname value="&zgname&"  disabled='disabled' >"& vbCrLf
	dwt.out"				  <INPUT type=hidden name=qxdj_zgname value="&zgname&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	if isnull(rsedit("zgdate")) then
	  zgdate=date()
	else
	  zgdate=rsedit("zgdate")  
	end if   
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>整改时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='qxdj_zgdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&zgdate&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

    if isnull(rsedit("qrname")) or rsedit("qrname")="" then 
	  qrname=session("UserName1")
	else
	  qrname=rsedit("qrname")  
	end if 
  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>确认人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_qrname value="&qrname&">"& vbCrLf
'	dwt.out"				  <INPUT type=hidden name=qxdj_qrname value="&qrname&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	
	if isnull(rsedit("qrdate")) then
	  qrdate=date()
	else
	  qrdate=rsedit("qrdate")  
	end if   
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>确认时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='qxdj_qrdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&qrdate&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	  
	
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='savefk'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	'sqlqxdj="SELECT * from qxdjzg ORDER BY id DESC"
	sqlqxdj="SELECT * from mzqxdjzg where not zgzt"
	if request("allchange")=1 then 	sqlqxdj="SELECT * from mzqxdjzg where zgzt"
	if keys<>"" then 
		sqlqxdj=sqlqxdj&" and wh like '%"&keys&"%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
          sqlqxdj=sqlqxdj&" and sscj="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	'if request("allnochange")=1 then sqlqxdj=sqlqxdj&" where zgjg=0"
	sqlqxdj=sqlqxdj&" ORDER BY jcdate1 desc"

	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>每周缺陷记录"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

'	for sscji=1 to 5 
'		sql="SELECT count(id) FROM mzqxdjzg WHERE sscj="&sscji&" and zgjg=0"
'		numb=numb&sscjh_d(sscji)&"<span style='color:#006600;'>"&connb.Execute(sql)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
'	next
	
'	sql="SELECT count(id) FROM mzqxdjzg WHERE  zgzt=false"
'	totall= "<span style='color:#006600;'>"&connb.Execute(sql)(0)&"</span>" 
'	dwt.out "<div class='pre'>未整改:"&numb&"合计:"&totall&"</div>"& vbCrLf
	call search()
	set rsqxdj=server.createobject("adodb.recordset")
	rsqxdj.open sqlqxdj,connb,1,1
	if rsqxdj.eof and rsqxdj.bof then 
	   message "未找到相关内容"
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>位号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>缺陷内容</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>检查日期</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>整改状态</div></td>"
		dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>整改结果</div></td>"
		'dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>整改</div></td>"
		dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>整改人</div></td>"
		dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>整改日期</div></td>"
		dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>确认人</div></td>"
		dwt.out "<td class='x-td'><DIV class='x-grid-hd-text'>确认日期</div></td>"

		
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
		dwt.out "    </tr>"
		record=rsqxdj.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
	   end if 
	   rsqxdj.PageSize = Cint(PgSz) 
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
	   rsqxdj.absolutePage = page
	   start=PgSz*Page-PgSz+1
	   rowCount = rsqxdj.PageSize
	   do while not rsqxdj.eof and rowcount>0
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&xh_id&"</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"
			dwt.out sscjh_d(rsqxdj("sscj"))&"</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"""
			'if now()-rsqxdj("update")<7 then   dwt.out "bgcolor=""#FFFF00"""
			dwt.out ">"
			if rsqxdj("zgzt") then 
			   dwt.out searchH(uCase(rsqxdj("wh")),keys)
			else
			   dwt.out "<font color='#ff0000'>"&searchH(uCase(rsqxdj("wh")),keys)&"<font>"
			end if  
			   dwt.out "&nbsp;</td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("body")&"&nbsp;</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("jcdate1")&"到"&rsqxdj("jcdate2")&"</div></td>"
			if rsqxdj("zgzt") then 
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">已整改</td>"
			else
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">未整改</td>"
			end if 
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("zgbody")&"&nbsp;</td>"
'			if rsqxdj("zgJG") then 
'			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">好</td>"
'			else
'			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">未好</td>"
'			end if 
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("zgname")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("zgdate")&"&nbsp;</div></td>"
			
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("qrname")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("qrdate")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"
		
			
			'if session("levelclass")=rsqxdj("sscj") and rsqxdj("zgzt")=false then dwt.out "反馈"
			if session("levelclass")=rsqxdj("sscj") then dwt.out "<a href=mzqxdjzg.asp?action=fk&id="&rsqxdj("id")&">反馈</a> "
			'if session("levelclass")=9 then dwt.out "<a href=qxdjzg.asp?action=edit&id="&rsqxdj("id")&"编辑</a> <a href=qxdjzg.asp?action=del&id="&rsqxdj("id")&"删除</a>"
			'dwt.out "<a href=mzqxdjzg.asp?action=edit&id="&rsqxdj("id")&">编辑</a> <a href=mzqxdjzg.asp?action=del&id="&rsqxdj("id")&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
			call editdel(rsqxdj("id"),rsqxdj("sscj"),"MZqxdjzg.asp?action=edit&id=","MZqxdjzg.asp?action=del&id=")
			dwt.out "</div></td></tr>"

			 RowCount=RowCount-1
          rsqxdj.movenext
		loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" or sscjid<>"" or request("allchange")=1then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
	rsqxdj.close
	set rsqxdj=nothing
	conn.close
	set conn=nothing
end sub
dwt.out "</body></html>"

sub search()
	dim sqlcj,rscj
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out " <form method='Get' name='SearchForm' action='qxdjzg.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "  <a href=""mzqxdjzg.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
	'dwt.out "  <a href=""mzqxdjzg.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
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
			dwt.out"<option value='mzqxdjzg.asp?sscj="&rscj("levelid")&"'"
	if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
			dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext	
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf
	dwt.out "<a href=?allchange=1>已整改记录</a> <a href=mzqxdjzg.asp>未整改记录</a>"
	dwt.out "</div></div></form>" & vbCrLf
end sub





Call CloseConn
%>