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
dim url,record,pgsz,total,page,start,rowcount,ii,pagename
dim keys,sscjid
'call conn_ndjx()
url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>信息管理系统年度检修管理页</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form1.ndjx_sscj.value==''){" & vbCrLf
Dwt.out "      alert('请选择所属车间！');" & vbCrLf
Dwt.out "   document.form1.ndjx_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form1.ndjx_wh.value==''){" & vbCrLf
Dwt.out "      alert('设备位号必须添写！');" & vbCrLf
Dwt.out "   document.form1.ndjx_wh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf


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
  case "addbj"
     call addbj
  case "saveaddbj"
    call saveaddbj
  case "editbj"
	call editbj
  case "saveeditbj"
    call saveeditbj
  case "delbj"
    call delbj



end select	




sub add()
dim sqlcj,rscj

	Dwt.out"<Div align=center><Div style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	Dwt.out"  <Div class=x-box-tl>"& vbCrLf
	Dwt.out"	<Div class=x-box-tr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-tc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-ml>"& vbCrLf
	Dwt.out"	<Div class=x-box-mr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-mc>"& vbCrLf
	Dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加年度检修</H3>"& vbCrLf
	Dwt.out"		<Div id=form-ct>"& vbCrLf
	Dwt.out "<form method='post' class='x-form' action='ndjx.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	Dwt.out"			<Div class='x-form-ct'>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>属所车间:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	if session("level")=0 then 
		Dwt.out"<div align=left><select name='ndjx_sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
		Dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
		sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
		Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
		rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		
		Dwt.out"</select></div>"  	 
	else 	 
		Dwt.out"<div align=left><input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></div></td></tr>"& vbCrLf
		Dwt.out"<input name='ndjx_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
	end if 
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>是否重点项目:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input type='checkbox' name='ndjx_iszd'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>设备位号:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_wh' type='text'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>数量:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_amount' type='text'  onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" ></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>检修内容:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 80px' name=ndjx_content >请添写检修内容</TEXTAREA></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-clear-left'></Div>"& vbCrLf
	
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>负责人:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_principal' type='text'></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>检修年度:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_nd' type='text'></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	
	
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>开工日期:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<div align=left><input name='ndjx_stardate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'></div>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>完工日期:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<div align=left><input name='ndjx_enddate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'></div>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"							<Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>备注:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				 <div align=left> <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=ndjx_bz></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	
	Dwt.out"			  <Div class=x-form-clear></Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"			<Div class=x-form-btns-ct>"& vbCrLf
	Dwt.out"			  <Div class='x-form-btns x-form-btns-center'>"& vbCrLf
	Dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	Dwt.out"				<Div class=x-clear></Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"		  </FORM>"& vbCrLf
	Dwt.out"		</Div>"& vbCrLf
	Dwt.out"	  </Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-bl>"& vbCrLf
	Dwt.out"	<Div class=x-box-br>"& vbCrLf
	Dwt.out"	  <Div class=x-box-bc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"</Div> "& vbCrLf  
	
   
   
end sub	

sub saveadd()  
dim sqladd,rsadd  
	 	on error resume next
 '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from ndjx_jx" 
      rsadd.open sqladd,connnd,1,3
      rsadd.addnew
	   rsadd("jx_sscj")=request("ndjx_sscj")
	   rsadd("jx_wh")=request("ndjx_wh")
	   rsadd("jx_amount")=request("ndjx_amount")
	   if request("ndjx_iszd")="on" then rsadd("jx_iszd")=true
	   rsadd("jx_principal")=request("ndjx_principal")
	   rsadd("jx_nd")=request("ndjx_nd")
	   rsadd("jx_stardate")=request("ndjx_stardate")
	   rsadd("jx_enddate")=request("ndjx_enddate")
	   rsadd("jx_content")=request("ndjx_content")
	   rsadd("jx_bz")=request("ndjx_bz")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub saveedit()  
dim rsedit,sqledit  
	  '保存
on error resume next
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from ndjx_jx where jx_id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connnd,1,3
      'rsedit("sscj")=Trim(Request("ndjx_sscj"))
	   'rsedit("jx_sscj")=request("ndjx_sscj")
	   rsedit("jx_wh")=request("ndjx_wh")
	   rsedit("jx_amount")=request("ndjx_amount")
	   if request("ndjx_iszd")="on" then rsedit("jx_iszd")=true
	   rsedit("jx_principal")=request("ndjx_principal")
	   rsedit("jx_stardate")=request("ndjx_stardate")
	   rsedit("jx_nd")=request("ndjx_nd")
	   rsedit("jx_enddate")=request("ndjx_enddate")
	   rsedit("jx_content")=request("ndjx_content")
	   rsedit("jx_bz")=request("ndjx_bz")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from ndjx_jx where jx_id="&id
  rsdel.open sqldel,connnd,1,3
  Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub

sub edit()
  	 

   
   dim sqledit,rsedit,id
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from ndjx_jx where jx_id="&id
   rsedit.open sqledit,connnd,1,1
   Dwt.out"<br><form method='post' action='ndjx.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   	Dwt.out"<Div align=center><Div style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	Dwt.out"  <Div class=x-box-tl>"& vbCrLf
	Dwt.out"	<Div class=x-box-tr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-tc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-ml>"& vbCrLf
	Dwt.out"	<Div class=x-box-mr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-mc>"& vbCrLf
	Dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑年度检修</H3>"& vbCrLf
	Dwt.out"		<Div id=form-ct>"& vbCrLf
   Dwt.out"<form method='post' action='ndjx.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	Dwt.out"			<Div class='x-form-ct'>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' >属所车间:</LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	
	Dwt.out"<div align=left><input class='x-form-text x-form-field' style='WIDTH: 175px'  value='"&sscjh(rsedit("jx_sscj"))&"'  disabled='disabled' ></div>"& vbCrLf
	dwt.out "<input type='hidden' name=ndjx_sscj value='"&sscjh(rsedit("jx_sscj"))&"'>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>是否重点项目:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input type='checkbox' name='ndjx_iszd' "
	if rsedit("jx_iszd") then dwt.out "CHECKED" 
	dwt.out"/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>设备位号:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_wh' type='text' value='"&rsedit("jx_wh")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>数量:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_amount' type='text'  value="&rsedit("jx_amount")&" onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" ></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>检修内容:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 80px' name=ndjx_content > "&rsedit("jx_content")&"</TEXTAREA></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-clear-left'></Div>"& vbCrLf
	
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>负责人:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_principal' type='text' value='"&rsedit("jx_principal")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>检修年度:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='ndjx_nd' type='text' value='"&rsedit("jx_nd")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	
	
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>开工日期:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<div align=left><input name='ndjx_stardate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("jx_stardate")&"'/></div>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>完工日期:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<div align=left><input name='ndjx_enddate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("jx_enddate")&"'/></div>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"							<Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px'><div align=right>备注:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				 <div align=left> <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=ndjx_bz value='"&rsedit("jx_bz")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	
	Dwt.out"			  <Div class=x-form-clear></Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"			<Div class=x-form-btns-ct>"& vbCrLf
	Dwt.out"			  <Div class='x-form-btns x-form-btns-center'>"& vbCrLf
	Dwt.out"			  <input name='action' type='hidden' value='saveedit'><input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	Dwt.out"				<Div class=x-clear></Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"		  </FORM>"& vbCrLf
	Dwt.out"		</Div>"& vbCrLf
	Dwt.out"	  </Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-bl>"& vbCrLf
	Dwt.out"	<Div class=x-box-br>"& vbCrLf
	Dwt.out"	  <Div class=x-box-bc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"</Div> "& vbCrLf 
	rsedit.close
	set rsedit=nothing 
end sub


sub addbj()
dim sqlcj,rscj

	Dwt.out"<Div align=center><Div style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	Dwt.out"  <Div class=x-box-tl>"& vbCrLf
	Dwt.out"	<Div class=x-box-tr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-tc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-ml>"& vbCrLf
	Dwt.out"	<Div class=x-box-mr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-mc>"& vbCrLf
	Dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加 "&request("sbwh")&" 备件</H3>"& vbCrLf
	Dwt.out"		<Div id=form-ct>"& vbCrLf
	Dwt.out "<form method='post' class='x-form' action='ndjx.asp' name='form1'>"
	Dwt.out"			<Div class='x-form-ct'>"& vbCrLf
				  
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>备件名称:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='name' type='text'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>规格型号:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='type' type='text'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>材质:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='cz' type='text'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>单位:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='dw' type='text'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>数量:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='numb' type='text'  onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" ></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>备注:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='bz' type='text'></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf
	
	
	
				  
	Dwt.out"			  <Div class=x-form-clear></Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"			<Div class=x-form-btns-ct>"& vbCrLf
	Dwt.out"			  <Div class='x-form-btns x-form-btns-center'>"& vbCrLf
	Dwt.out"			  <input name='action' type='hidden' value='saveaddbj'>  <input name='jxid' type='hidden' value='"&request("jxid")&"'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	Dwt.out"				<Div class=x-clear></Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"		  </FORM>"& vbCrLf

	Dwt.out"		</Div>"& vbCrLf
	Dwt.out"	  </Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-bl>"& vbCrLf
	Dwt.out"	<Div class=x-box-br>"& vbCrLf
	Dwt.out"	  <Div class=x-box-bc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"</Div> "& vbCrLf  
	
   
   
end sub	



sub editbj()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from ndjx_bj where id="&id
   rsedit.open sqledit,connnd,1,1

	Dwt.out"<Div align=center><Div style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	Dwt.out"  <Div class=x-box-tl>"& vbCrLf
	Dwt.out"	<Div class=x-box-tr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-tc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-ml>"& vbCrLf
	Dwt.out"	<Div class=x-box-mr>"& vbCrLf
	Dwt.out"	  <Div class=x-box-mc>"& vbCrLf
	Dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑 "&request("sbwh")&" 备件</H3>"& vbCrLf
	Dwt.out"		<Div id=form-ct>"& vbCrLf
	Dwt.out "<form method='post' class='x-form' action='ndjx.asp' name='form1'>"
	Dwt.out"			<Div class='x-form-ct'>"& vbCrLf
				  
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>备件名称:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='name' type='text' value='"&rsedit("name")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>规格型号:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='type' type='text' value='"&rsedit("type")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>材质:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='cz' type='text' value='"&rsedit("cz")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>单位:</div></LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='dw' type='text' value='"&rsedit("dw")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
				  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>数量:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='numb' type='text'  value='"&rsedit("numb")&"' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;""/ ></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf


	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' ><div align=right>备注:</LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out"				  <div align=left><input name='bz' type='text' value='"&rsedit("bz")&"'/></div>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
	Dwt.out"							<Div class='x-form-clear-left'></Div>"& vbCrLf
	
	
	
				  
	Dwt.out"			  <Div class=x-form-clear></Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"			<Div class=x-form-btns-ct>"& vbCrLf
	Dwt.out"			  <Div class='x-form-btns x-form-btns-center'>"& vbCrLf
	Dwt.out"			  <input name='action' type='hidden' value='saveeditbj'>  <input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	Dwt.out"				<Div class=x-clear></Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			</Div>"& vbCrLf
	Dwt.out"		  </FORM>"& vbCrLf

	Dwt.out"		</Div>"& vbCrLf
	Dwt.out"	  </Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"  <Div class=x-box-bl>"& vbCrLf
	Dwt.out"	<Div class=x-box-br>"& vbCrLf
	Dwt.out"	  <Div class=x-box-bc></Div>"& vbCrLf
	Dwt.out"	</Div>"& vbCrLf
	Dwt.out"  </Div>"& vbCrLf
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"</Div> "& vbCrLf  
	rsedit.close
	set rsedit=nothing 
end sub	

sub saveaddbj()  
	 	on error resume next
 '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from ndjx_bj" 
      rsadd.open sqladd,connnd,1,3
      rsadd.addnew
	   rsadd("jx_id")=request("jxid")
	   rsadd("name")=request("name")
	   rsadd("type")=request("type")
	   rsadd("cz")=request("cz")
	   rsadd("numb")=request("numb")
	   rsadd("dw")=request("dw")
	   rsadd("bz")=request("bz")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub
sub saveeditbj()  
	 	'on error resume next
 '保存
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from ndjx_bj" 
      rsedit.open sqledit,connnd,1,3
	   'rsedit("jx_id")=request("jxid")
	   rsedit("name")=request("name")
	   rsedit("type")=request("type")
	   rsedit("cz")=request("cz")
	   rsedit("numb")=request("numb")
	   rsedit("dw")=request("dw")
	   rsedit("bz")=request("bz")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub
sub delbj()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from ndjx_bj where id="&id
  rsdel.open sqldel,connnd,1,3
  Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub

sub main()
	Dwt.Out "<SCRIPT language=javascript1.2>" & vbCrLf
	Dwt.Out "function showsubmenu(sid){" & vbCrLf
	Dwt.Out "      	 var ss='xxx'+sid;" & vbCrLf
	Dwt.Out "    whichEl = eval('info' + sid);" & vbCrLf
	Dwt.Out "    if (whichEl.style.display == 'none'){" & vbCrLf
	Dwt.Out "        eval(""info"" + sid + "".style.display='block';"");" & vbCrLf
	Dwt.Out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i6.gif' />"";" & vbCrLf
	Dwt.Out "    }" & vbCrLf
	Dwt.Out "    else{" & vbCrLf
	Dwt.Out "        eval(""info"" + sid + "".style.display='none';"");" & vbCrLf
	Dwt.Out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i7.gif' />"";" & vbCrLf
	Dwt.Out "    }" & vbCrLf
	Dwt.Out "}" & vbCrLf
	Dwt.Out "</SCRIPT>" & vbCrLf

	
	dim sqlndjx,rsndjx,title
	sqlndjx="SELECT * from ndjx_jx where 1=1"
	'if keys<>"" or sscjid<>"" or request("jx_nd")<>"" then sqlndjx=sqlndjx&" where "

	if keys<>"" then 
		sqlndjx=sqlndjx&" and jx_wh like '%" &keys& "%' "
		title=title&"-搜索 "&keys
	end if 
	if request("jx_nd")<>"" then
		sqlndjx=sqlndjx&" and jx_nd="&request("jx_nd")
		title=title&"-"&request("jx_nd")&"年"
	end if 
	if sscjid<>"" then
		sqlndjx=sqlndjx&" and jx_sscj="&sscjid
		title=title&"-"&sscjh(sscjid)
	end if 
	sqlndjx=sqlndjx&" ORDER BY jx_sscj aSC,jx_stardate desc"
	
	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>年度检修"&title&"</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
	
'	'for sscji=1 to 5 '071017修改
'	sql="select * from levelname where istq=false"
'	set rs=server.createobject("adodb.recordset")
'	rs.open sql,conn,1,1
'	if rs.eof and rs.bof then 
'		Dwt.out "没有添加车间"
'	else
'	   do while not rs.eof
'		sql="SELECT count(id) FROM ndjx WHERE sscj="&rs("levelid")&" and month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
'		numb=numb&sscjh_d(rs("levelid"))&":"&"<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
'	rs.movenext
'	loop
'	end if 
'	rs.close
'	
'	sql="SELECT count(id) FROM ndjx WHERE  month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
'	totall= "<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>" 
'	Dwt.out "<Div class='pre'>本月"&numb&"合计:"&totall&"</Div>"& vbCrLf

	search()
	
	set rsndjx=server.createobject("adodb.recordset")
	rsndjx.open sqlndjx,connndjx,1,1
	if rsndjx.eof and rsndjx.bof then 
		message("未找到相关内容")
	else
		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.out "     <td class='x-td'><Div class='x-grid-hd-text'>序号</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>车间</Div></td>"& vbCrLf
		'Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>项目编号</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>设备位号名称</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>数量</Div></td>"& vbCrLf
		'Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>检修内容</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>负责人</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>检修年度</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>开工日期</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>完工日期</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>备注</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>选项</Div></td>"& vbCrLf
		Dwt.out "    </tr>"& vbCrLf
		record=rsndjx.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rsndjx.PageSize = Cint(PgSz) 
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
		rsndjx.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rsndjx.PageSize
		do while not rsndjx.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center""><a href='#' onclick=""showsubmenu("&rsndjx("jx_id")&");"" id=xxx"&rsndjx("jx_id")&"><img src='/img_ext/i7.gif' /></a>"&xh_id&"</Div></td>"& vbCrLf
			Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&sscjh(rsndjx("jx_sscj"))&"</Div></td>"& vbCrLf
			'Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_number")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"
			if rsndjx("jx_iszd") then 
			  dwt.out "<span style='color:#ff0000'>"&searchH(uCase(rsndjx("jx_wh")),keys)&"</span>"
			else
              dwt.out searchH(uCase(rsndjx("jx_wh")),keys)
			end if 
			dwt.out "</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsndjx("jx_amount")&"</Div></td>"& vbCrLf
			'Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_content")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_principal")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_nd")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_stardate")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_enddate")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsndjx("jx_bz")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=center>"& vbCrLf
			call editdel(rsndjx("jx_id"),rsndjx("jx_sscj"),"ndjx.asp?action=edit&id=","ndjx.asp?action=del&id=")
			Dwt.out "</Div></td>"
			Dwt.OUT "</tr>"& vbCrLf
	
	        sqlbj="SELECT * from ndjx_bj where jx_id="&rsndjx("jx_id")
			set rsbj=server.createobject("adodb.recordset")
			rsbj.open sqlbj,connndjx,1,1
			if rsbj.eof and rsbj.bof then 
				Dwt.Out "<tr class='x-grid-row'><td    bgcolor='#BFDFFF' colspan=7 style='display:none' id='info"&rsndjx("jx_id")&"'>"		
				if session("levelclass")=rsndjx("jx_sscj") then dwt.out "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=ndjx.asp?action=addbj&jxid="&rsndjx("jx_id")&"&sbwh='"&uCase(rsndjx("jx_wh"))&"'>添加备件</a>"
				dwt.out "<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td bgcolor='#BFDFFF' width='20%'>检修内容</td><td bgcolor='#BFDFFF'>"&rsndjx("jx_content")&"</td></tr></table>"
				dwt.out "</td></tr>"
			else
				Dwt.Out "<tr class='x-grid-row'><td  colspan=7 style='display:none' id='info"&rsndjx("jx_id")&"'>"		
				if session("levelclass")=rsndjx("jx_sscj") then dwt.out "<a href=ndjx.asp?action=addbj&jxid="&rsndjx("jx_id")&"&sbwh='"&uCase(rsndjx("jx_wh"))&"'>添加备件</a>"
				dwt.out "<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td bgcolor='#BFDFFF' width='20%'>检修内容</td><td bgcolor='#BFDFFF'>"&rsndjx("jx_content")&"</td></tr></table>"
				Dwt.Out "<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
				Dwt.Out "<tr >" & vbCrLf
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>备件名称</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>规格型号</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>材质</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>单位</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>数量</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>选项</Div></td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>选项</Div></td>"
				Dwt.Out  "    </tr>"
			do while not rsbj.eof
				Dwt.Out "<tr class='x-grid-row'  >"& vbCrLf
				Dwt.Out "      <td   bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("name")&"&nbsp;</td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("type")&"&nbsp;</td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("cz")&"&nbsp;</td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("dw")&"&nbsp;</td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("numb")&"&nbsp;</td>"
				Dwt.Out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rsbj("bz")&"&nbsp;</td>"
				Dwt.out "      <td   bgcolor='#BFDFFF'  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=center>"& vbCrLf
				call editdel(rsbj("id"),rsndjx("jx_sscj"),"ndjx.asp?action=editbj&sbwh="&rsndjx("jx_wh")&"&id=","ndjx.asp?action=delbj&id=")
				Dwt.out "</Div></td>"
				Dwt.Out  "    </tr>"
					rsbj.movenext
				loop
				Dwt.Out "</table>"		
				Dwt.Out "</td></tr>"		
			end if 
			RowCount=RowCount-1
			rsndjx.movenext
		loop
		Dwt.out "</table>"& vbCrLf
		if keys<>"" or sscjid<>"" or request("jx_nd")<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 
		Dwt.out "</Div>"& vbCrLf
	end if
	Dwt.out "</Div>"  
	rsndjx.close
	set rsndjx=nothing
	conn.close
	set conn=nothing
end sub

Dwt.out "</body></html>"

sub search()
	dim sqlcj,rscj
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='ndjx.asp'>" & vbCrLf
	
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.out "<a href=""ndjx.asp?action=add"">添加年度检修</a>&nbsp;&nbsp;"
	
	Dwt.out "  <input type='text' name='keyword'  size='20' maxlength='50' "
	if keys<>"" then 
		 Dwt.out "value='"&keys&"'"
    	Dwt.out ">" & vbCrLf
    else
		 Dwt.out "value='输入搜索的设备位号'"
	 	Dwt.out " onblur=""if(this.value==''){this.value='输入搜索的设备位号'}"" onfocus=""this.value=''"">" & vbCrLf
	end if                 
	Dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	
	
	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlgh="SELECT distinct jx_sscj from ndjx_jx"
	if request("jx_nd")<>"" then sqlgh=sqlgh&" where jx_nd="&request("jx_nd")
    sqlgh=sqlgh&" order by jx_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connndjx,1,1
    do while not rsgh.eof
		cjid=cint(rsgh("jx_sscj"))
		sql="SELECT count(jx_id) FROM ndjx_jx WHERE jx_sscj="&cjid
		if request("jx_nd")<>"" then sql=sql&" and jx_nd="&request("jx_nd")
		jx_numb=connnd.Execute(sql)(0)
        
		if jx_numb<>0 then 
			'i=i+1
			Dwt.out"<option  value='ndjx.asp?sscj="&cjid
		    if request("jx_nd")<>"" then dwt.out "&jx_nd="&request("jx_nd")
			dwt.out "'"
			if cint(request("sscj"))=cjid then Dwt.out" selected"


			sql="SELECT levelname FROM levelname WHERE levelid="&cjid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    cj_name="未知项"
			else
			    cj_name=rs("levelname")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&cj_name&"("&jx_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf

	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按年度跳转至…</option>" & vbCrLf
	sqlgh="SELECT distinct jx_nd from ndjx_jx"
	if request("sscj")<>"" then sqlgh=sqlgh&" where jx_sscj="&request("sscj")
    'sqlgh=sqlgh&" order by jx_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connndjx,1,1
    do while not rsgh.eof
		jx_nd=cint(rsgh("jx_nd"))
		sql="SELECT count(jx_id) FROM ndjx_jx WHERE jx_nd="&jx_nd
		if request("sscj")<>"" then sql=sql&" and jx_sscj="&request("sscj")
		jx_numb=connnd.Execute(sql)(0)
        
		if jx_numb<>0 then 
			'i=i+1
			Dwt.out"<option  value='ndjx.asp?jx_nd="&jx_nd&"'"
			if cint(request("jx_nd"))=jx_nd then Dwt.out" selected"
			Dwt.out ">"&jx_nd&"("&jx_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf
'	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
'	Dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'	set rscj=server.createobject("adodb.recordset")
'	rscj.open sqlcj,conn,1,1
'	do while not rscj.eof
'		Dwt.out"<option value='ndjx.asp?sscj="&rscj("levelid")&"'"
'		if cint(request("sscj"))=rscj("levelid") then Dwt.out" selected"
'		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'	Dwt.out "</select>&nbsp;&nbsp;"
	'dwt.out"<a href=tocsv.asp?action=dcsjxmain&sql1=ndjx&titlename=检修记录>输出检修记录到Excel文档</a>	" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf
end sub





Call CloseConn
%>