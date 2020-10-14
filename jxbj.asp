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
dwt.out "<title>信息管理系统-年度检修计划备件汇总</title>"& vbCrLf
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
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case "isck"
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from qxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connb,1,3
		rsedit("isck")=true
	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
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
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加年度检修备件</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='jxbj.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' >检修年度:</LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out "<select name='jx_nd' style='WIDTH: 175px'/>"
	for  i=year(now())-3 to year(now())+3
         Dwt.out "<option value="&i
		 if i=year(now()) then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	if session("level")=0 or  session("groupid")=7 then 
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
	dwt.out"				<LABEL style='WIDTH: 85px' >名称:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=name>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px' >型号:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				   <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=type>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>材质:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=cz>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>单位:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=dw>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>数量:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=numb>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>定货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='dhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>计划到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jhdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>实际到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='sjdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>备注:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=bz>"& vbCrLf
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
	sqladd="select * from ndjx_bj" 
	rsadd.open sqladd,connnd,1,3
	rsadd.addnew
	on error resume next
	rsadd("jx_nd")=request("jx_nd")
	rsadd("sscj")=Trim(Request("sscj"))
	rsadd("name")=request("name")
	rsadd("type")=Trim(request("type"))
	rsadd("cz")=request("cz")
	rsadd("dw")=request("dw")
	rsadd("numb")=request("numb")
	rsadd("bz")=request("bz")
	rsadd("dhdate")=request("dhdate")
	rsadd("jhdhdate")=request("jhdhdate")
	rsadd("sjdhdate")=request("sjdhdate")
	rsadd.update
	rsadd.close
	set rsadd=nothing
	dwt.out"<Script Language=Javascript>location.href='jxbj.asp';</Script>"
end sub


sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from ndjx_bj where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connnd,1,3
	on error resume next
	'rsedit("sscj")=Trim(Request("qxdj_sscj"))
	rsedit("jx_nd")=request("jx_nd")
	rsedit("sscj")=Trim(Request("sscj"))
	rsedit("name")=request("name")
	rsedit("type")=Trim(request("type"))
	rsedit("cz")=request("cz")
	rsedit("dw")=request("dw")
	rsedit("numb")=request("numb")
	rsedit("bz")=request("bz")
	rsedit("dhdate")=request("dhdate")
	rsedit("jhdhdate")=request("jhdhdate")
	rsedit("sjdhdate")=request("sjdhdate")
	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
	qxdjid=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from ndjx_bj where id="&qxdjid
	rsdel.open sqldel,connnd,1,3
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub


sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from ndjx_bj where id="&id
	rsedit.open sqledit,connnd,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑大修备件</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='jxbj.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 85px' >检修年度:</LABEL>"& vbCrLf
	Dwt.out"				<Div class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out "<select name='jx_nd' style='WIDTH: 175px'/>"
	for  i=year(now())-3 to year(now())+3
         Dwt.out "<option value="&i
		 if i=cint(rsedit("jx_nd")) then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	if session("level")=0 or  session("groupid")=7 then 
		dwt.out"<select name='sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  value='"&rsedit("sscj")&"' selected>"&sscjh(rsedit("sscj"))&"</option>"& vbCrLf
		sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
		if rsedit("sscj")<>rscj("levelid") then
		dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
		end if
		rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		
		dwt.out"</select>"  	 
	else 	 
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
		dwt.out"<input name='sscj' type='hidden' value='"&sscjh(rsedit("sscj"))&"></td></tr>"& vbCrLf
	end if 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px' >名称:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=name value='"&rsedit("name")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px' >型号:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				   <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=type value='"&rsedit("type")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>材质:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=cz value='"&rsedit("cz")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>单位:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=dw value='"&rsedit("dw")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>数量:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=numb value='"&rsedit("numb")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>定货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='dhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("dhdate")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>计划到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jhdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly value='"&rsedit("jhdhdate")&"' >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>实际到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='sjdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly value='"&rsedit("sjdhdate")&"' >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>备注:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=bz value='"&rsedit("bz")&"'>"& vbCrLf
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
	'sqlqxdj="SELECT * from qxdjzg ORDER BY id DESC"
	sqlqxdj="SELECT * from ndjx_bj  where 1=1"

	if keys<>"" then 
		sqlqxdj=sqlqxdj&" and name like '%" &keys& "%' "
		title=title&"-搜索 "&keys
	end if 
	if request("jx_nd")<>"" then
		sqlqxdj=sqlqxdj&" and jx_nd="&request("jx_nd")
		title=title&"-"&request("jx_nd")&"年"
	end if 
	if sscjid<>"" then
		sqlqxdj=sqlqxdj&" and sscj="&sscjid
		title=title&"-"&sscjh(sscjid)
	end if 
	
	
	'if request("allnochange")=1 then sqlqxdj=sqlqxdj&" where zgjg=0"
	sqlqxdj=sqlqxdj&" ORDER BY sscj deSC,jx_id desc,id desc"

	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>年度检修备件汇总"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

	
	call search()
	set rsqxdj=server.createobject("adodb.recordset")
	rsqxdj.open sqlqxdj,connnd,1,1
	if rsqxdj.eof and rsqxdj.bof then 
	   message "未找到相关内容"
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>名称</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>型号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>材质</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>单位</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>数量</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>订货</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>计划到货</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>实际到货</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>备注</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>大修年度</div></td>"
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
			if rsqxdj("sscj")<>"" then dwt.out sscjh_d(rsqxdj("sscj"))
			
			dwt.out"</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&searchH(uCase(rsqxdj("name")),keys)&"</td>"'
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("type")&"&nbsp;</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("cz")&"&nbsp;</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("dw")&"&nbsp;</td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("numb")&"&nbsp;</div></td>"
			dhdate=rsqxdj("dhdate")
			if dhdate="" or isnull(dhdate) then 
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未定货&nbsp;</span></div></td>"
			else
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate&"&nbsp;</div></td>"
			end if
						
			dhdate1=rsqxdj("jhdhdate")
			if dhdate1="" or isnull(dhdate1) then 
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未定货&nbsp;</span></div></td>"
			else
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate1&"&nbsp;</div></td>"
			end if

			dhdate2=rsqxdj("sjdhdate")
			if dhdate2="" or isnull(dhdate2) then 
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未定货&nbsp;</span></div></td>"
			else
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate2&"&nbsp;</div></td>"
			end if
						
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("bz")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("jx_nd")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"
			'如果LEVELCLASS=7则显示设置厂控缺陷
			'call editdel(rsqxdj("id"),rsqxdj("sscj"),"jxbj.asp?action=edit&id=","jxbj.asp?action=del&id=")
            if session("groupid")=7 then
 
    dwt.out "<a href=jxbj.asp?action=edit&id="&rsqxdj("id")&">编辑</a>&nbsp;"
    dwt.out "<a href=jxbj.asp?action=del&id="&rsqxdj("id")&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 end if 
 dwt.out "&nbsp;"

			dwt.out "</div></td></tr>"
			 RowCount=RowCount-1
          rsqxdj.movenext
		loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" or sscjid<>"" or request("allnochange")=1 or request("jx_nd")<>"" then
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
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='jxbj.asp'>" & vbCrLf
	
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.out "<a href=""jxbj.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
	
	Dwt.out "  <input type='text' name='keyword'  size='20' maxlength='50' "
	if keys<>"" then 
		 Dwt.out "value='"&keys&"'"
    	Dwt.out ">" & vbCrLf
    else
		 Dwt.out "value='输入名称'"
	 	Dwt.out " onblur=""if(this.value==''){this.value='输入名称'}"" onfocus=""this.value=''"">" & vbCrLf
	end if                 
	Dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	
	
	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlgh="SELECT distinct sscj from ndjx_bj"
	if request("jx_nd")<>"" then sqlgh=sqlgh&" where jx_nd="&request("jx_nd")
    sqlgh=sqlgh&" order by sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connndjx,1,1
    do while not rsgh.eof
		cjid=cint(rsgh("sscj"))
		sql="SELECT count(jx_id) FROM ndjx_bj WHERE sscj="&cjid
		if request("jx_nd")<>"" then sql=sql&" and jx_nd="&request("jx_nd")
		jx_numb=connnd.Execute(sql)(0)
        
		if jx_numb<>0 then 
			'i=i+1
			Dwt.out"<option  value='jxbj.asp?sscj="&cjid
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
	sqlgh="SELECT distinct jx_nd from ndjx_bj"
	if request("sscj")<>"" then sqlgh=sqlgh&" where sscj="&request("sscj")
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connndjx,1,1
    do while not rsgh.eof
		jx_nd=cint(rsgh("jx_nd"))
		sql="SELECT count(id) FROM ndjx_bj WHERE jx_nd="&jx_nd
		if request("sscj")<>"" then sql=sql&" and sscj="&request("sscj")
		jx_numb=connnd.Execute(sql)(0)
        
		if jx_numb<>0 then 
			i=i+1
			Dwt.out"<option  value='jxbj.asp?jx_nd="&jx_nd&"'"
			if cint(request("jx_nd"))=jx_nd then Dwt.out" selected"
			Dwt.out ">"&jx_nd&"("&jx_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf

end sub

	



Call CloseConn
%>