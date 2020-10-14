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
<!--#include file="inc/function.asp"-->
<!--#include file="inc/session.asp"-->

<%
dim title,record,pgsz,total,page,start,rowcount,url,ii,zxzz
dim id,scontent,rsdel,sqldel
url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>外检计量器具台账</title>"& vbCrLf
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
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='jlqjtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>属所单位:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='ssdw' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected>选择所属单位</option>"& vbCrLf

                dwt.out"<option value=23>11</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>计量器具编号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='bh' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>计量器具名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='name' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>规格型号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ggxh' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>精度等级:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='jddj' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>测量范围:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='clfw' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>制造厂家:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='zzcj' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>出厂编号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ccbh' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>使用地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='sydd' type='text' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>管理方式:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='glfs' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected>选择管理方式</option>"& vbCrLf
		dwt.out"<option value=1>A</option>"& vbCrLf
		dwt.out"<option value=2>B</option>"& vbCrLf
		dwt.out"<option value=3>C</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='jddd' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected>选择检定地点</option>"& vbCrLf
		dwt.out"<option value=1>中国计量局</option>"& vbCrLf
		dwt.out"<option value=2> 省质监局</option>"& vbCrLf
		dwt.out"<option value=3> 市计量局</option>"& vbCrLf
 		dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定周期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='jdzq' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected>选择检定周期</option>"& vbCrLf
		dwt.out"<option value=1>6个月</option>"& vbCrLf
		dwt.out"<option value=2>12个月</option>"& vbCrLf
		dwt.out"<option value=3>24个月</option>"& vbCrLf
		dwt.out"<option value=4>48个月</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定计划:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='jdjh' type='text' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>启用时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='qydate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)'  >"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>参考价:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ckj' type='text' >"& vbCrLf
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
	sqladd="select * from jlqjtz" 
	rsadd.open sqladd,connb,1,3
	rsadd.addnew
	on error resume next
	rsadd("ssdw")=request("ssdw")
	rsadd("bh")=request("bh")
	rsadd("name")=request("name")
	rsadd("ggxh")=request("ggxh")
	rsadd("jddj")=request("jddj")
	rsadd("clfw")=request("clfw")
	rsadd("zzcj")=request("zzcj")
	rsadd("qydate")=request("qydate")
	rsadd("sydd")=request("sydd")
	rsadd("glfs")=request("glfs")
	rsadd("jddd")=request("jddd")
	rsadd("jdzq")=request("jdzq")
	rsadd("jdjh")=request("jdjh")
	rsadd("ccbh")=request("ccbh")
	rsadd("ckj")=request("ckj")
	rsadd("bz")=request("bz")
	
	rsadd.update
	rsadd.close
	Dwt.savesl "计量器具台账","编辑",request("name")
	set rsadd=nothing
	dwt.out"<Script Language=Javascript>location.href='jlqjtz.asp';</Script>"
end sub


sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from jlqjtz where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connb,1,3
	'on error resume next
	rsedit("ssdw")=request("ssdw")
	rsedit("bh")=request("bh")
	rsedit("name")=request("name")
	rsedit("ggxh")=request("ggxh")
	rsedit("jddj")=request("jddj")
	rsedit("clfw")=request("clfw")
	rsedit("zzcj")=request("zzcj")
	rsedit("qydate")=request("qydate")
	rsedit("sydd")=request("sydd")
	rsedit("glfs")=request("glfs")
	rsedit("jddd")=request("jddd")
	rsedit("jdzq")=request("jdzq")
	rsedit("jdjh")=request("jdjh")
	rsedit("ccbh")=request("ccbh")
	rsedit("ckj")=request("ckj")
	rsedit("bz")=request("bz")

	rsedit.update
	rsedit.close
	Dwt.savesl "计量器具台账","编辑",request("name")
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
	id=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from JLQJtz where id="&id
	rsdel.open sqldel,connb,1,3
	Dwt.savesl "计量器具台账","删除",id
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub


sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from jlqjtz where id="&id
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
    dwt.out"<form method='post' action='jlqjtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' >属所单位:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='ssdw' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option value=1"
		if rsedit("ssdw")=1 then dwt.out " selected"
		dwt.out ">质监处</option>"& vbCrLf
		dwt.out"<option value=2"
		if rsedit("ssdw")=2 then dwt.out " selected"
		dwt.out">合成厂</option>"& vbCrLf
		dwt.out"<option value=3"
		if rsedit("ssdw")=3 then dwt.out " selected"
		dwt.out">复肥厂</option>"& vbCrLf
		dwt.out"<option value=4"
		if rsedit("ssdw")=4 then dwt.out " selected"
		dwt.out ">热动厂</option>"& vbCrLf
		dwt.out"<option value=5"
		if rsedit("ssdw")=5 then dwt.out " selected"
		dwt.out"></option>"& vbCrLf
		dwt.out"<option value=6"
		if rsedit("ssdw")=6 then dwt.out " selected"
		dwt.out">电气厂</option>"& vbCrLf
		dwt.out"<option value=7"
		if rsedit("ssdw")=7 then dwt.out " selected"
        dwt.out ">供水厂</option>"& vbCrLf
		dwt.out"<option value=8"
		if rsedit("ssdw")=8 then dwt.out " selected"
		dwt.out">备煤厂</option>"& vbCrLf
		dwt.out"<option value=9"
		if rsedit("ssdw")=9 then dwt.out " selected"
		dwt.out">成品厂</option>"& vbCrLf
		dwt.out"<option value=10"
		if rsedit("ssdw")=10 then dwt.out " selected"
		dwt.out">硝铵</option>"& vbCrLf
		dwt.out"<option value=11"
		if rsedit("ssdw")=11 then dwt.out " selected"
		dwt.out">硝酸</option>"& vbCrLf
		dwt.out"<option value=12"
		if rsedit("ssdw")=12 then dwt.out " selected"
		dwt.out">消防队</option>"& vbCrLf
		dwt.out"<option value=13"
		if rsedit("ssdw")=13 then dwt.out " selected"
		dwt.out">苯胺</option>"& vbCrLf
		dwt.out"<option value=14"
		if rsedit("ssdw")=14 then dwt.out " selected"
		dwt.out">环保处</option>"& vbCrLf
		dwt.out"<option value=15"
		if rsedit("ssdw")=15 then dwt.out " selected"
		dwt.out"></option>"& vbCrLf
		dwt.out"<option value=16"
		if rsedit("ssdw")=16 then dwt.out " selected"
		dwt.out">建筑公司</option>"& vbCrLf
		dwt.out"<option value=17"
		if rsedit("ssdw")=17 then dwt.out " selected"
		dwt.out">水泥厂</option>"& vbCrLf
		dwt.out"<option value=18"
		if rsedit("ssdw")=18 then dwt.out " selected"
		dwt.out">精细公司</option>"& vbCrLf
		dwt.out"<option value=19"
		if rsedit("ssdw")=19 then dwt.out " selected"
		dwt.out">钾盐公司</option>"& vbCrLf
		dwt.out"<option value=20"
		if rsedit("ssdw")=20 then dwt.out " selected"
		dwt.out">服务公司</option>"& vbCrLf
		dwt.out"<option value=21"
		if rsedit("ssdw")=21 then dwt.out " selected"
		dwt.out">应化公司</option>"& vbCrLf
		dwt.out"<option value=22"
		if rsedit("ssdw")=22 then dwt.out " selected"
		dwt.out">塑料公司</option>"& vbCrLf
                dwt.out"<option value=23"
		if rsedit("ssdw")=23 then dwt.out " selected"
		dwt.out">供销公司</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>计量器具编号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='bh' type='text' value="&rsedit("bh")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>计量器具名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='name' type='text' value="&rsedit("name")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>规格型号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ggxh' type='text' value="&rsedit("ggxh")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>精度等级:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='jddj' type='text' value="&rsedit("jddj")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>测量范围:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='clfw' type='text' value="&rsedit("clfw")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>制造厂家:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='zzcj' type='text' value="&rsedit("zzcj")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>出厂编号:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ccbh' type='text' value="&rsedit("ccbh")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>使用地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='sydd' type='text'  value="&rsedit("sydd")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>管理方式:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='glfs' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option value=1"
		if rsedit("glfs")=1 then dwt.out " selected"
		dwt.out">A</option>"& vbCrLf
		dwt.out"<option value=2"
		if rsedit("glfs")=2 then dwt.out " selected"
		dwt.out">B</option>"& vbCrLf
		dwt.out"<option value=3"
		if rsedit("glfs")=3 then dwt.out " selected"
		dwt.out">C</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定地点:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='jddd' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option value=1"
		if rsedit("jddd")=1 then dwt.out " selected"
		dwt.out">中国计量局</option>"& vbCrLf
		dwt.out"<option value=2"
		if rsedit("jddd")=2 then dwt.out " selected"
		dwt.out"> 省质监局</option>"& vbCrLf
		dwt.out"<option value=3"
		if rsedit("jddd")=3 then dwt.out " selected"
		dwt.out"> 市计量局</option>"& vbCrLf
		dwt.out"<option value=4"
		if rsedit("jddd")=4 then dwt.out " selected"
		dwt.out"> 计量局</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定周期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
		dwt.out"<select name='jdzq' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option value=1"
		if rsedit("jdzq")=1 then dwt.out " selected"
		dwt.out">6个月</option>"& vbCrLf
		dwt.out"<option value=2"
		if rsedit("jdzq")=2 then dwt.out " selected"
		dwt.out">12个月</option>"& vbCrLf
		dwt.out"<option value=3"
		if rsedit("jdzq")=3 then dwt.out " selected"
		dwt.out">24个月</option>"& vbCrLf
		dwt.out"<option value=4"
		if rsedit("jdzq")=4 then dwt.out " selected"
		dwt.out">48个月</option>"& vbCrLf
		dwt.out"</select>"  	 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>检定计划:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='jdjh' type='text'  value="&rsedit("jdjh")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>启用时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='qydate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)'   value="&rsedit("qydate")&">"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>参考价:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='ckj' type='text'  value="&rsedit("ckj")&">"& vbCrLf
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
	sql="SELECT * from jlqjtz"
	if keys<>"" then 
		sql=sql&" where name like '%"&keys&"%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
        sql=sql&" where ssdw="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	sql=sql&" ORDER BY bh aSC"

	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>外检计量器具台账"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

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
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>所属单位</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>编号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>名称</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>规格型号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>精度等级</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>范围</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>厂家</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>出厂编号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>启用日期</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>使用地点</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>管理方式</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>检定地点</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>周期</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>备注</div></td>"
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
			dwt.out "     <td  class='tdcl tdal'>"&xh_id&"</td>"
			dwt.out "      <td class='tdcl tdal'>"
			select case rs("ssdw")
				case 1
				    ssdw="质监处"
				case 2
				ssdw="合成厂"
				case 3
				ssdw="复肥厂"
				case 4
				ssdw="热动厂"
				case 5
				ssdw=""
				case 6
				ssdw="电气厂"
				case 7
				ssdw="供水厂"
				case 8
				ssdw="备煤厂"
				case 9
				ssdw="成品厂"
				case 10
				ssdw="硝铵"
				case 11
				ssdw="硝酸"
				case 12
				ssdw="消防队"
				case 13
				ssdw="苯胺"
				case 14
				ssdw="环保处"
				case 15
				ssdw=""
				case 16
				ssdw="建筑公司"
				case 17
				ssdw="水泥厂"
				case 18
				ssdw="精细公司"
				case 19
				ssdw="钾盐公司"
				case 20
				ssdw="服务公司"
				case 21
				ssdw="应化公司"
				case 22
				ssdw="塑料公司"
                                case 23
				ssdw="供销公司"
			end select
			dwt.out ssdw	
		dwt.out"</select>"  	 

			dwt.out"</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("bh")&"</td>"
			dwt.out "      <td class='tdcl'>"&rs("name")&"</td>"
			dwt.out "      <td class='tdcl tdbr' >"&rs("ggxh")&"</td>"
			dwt.out "      <td class='tdcl tdbr tdal' >"&rs("jddj")&"</td>"
			dwt.out "      <td class='tdcl'>"&rs("clfw")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("zzcj")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("ccbh")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("qydate")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("sydd")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"
			select case rs("glfs")
			    case 1
				 	glfs="A"
				case 2
				    glfs="B"
		        case 3
				    glfs="C"
			end select
			
			dwt.out glfs				 

			
			dwt.out"</td>"
			dwt.out "      <td class='tdcl tdal'>"
			select case rs("jddd")
			    case 1
				 	jddd="中国计量局"
				case 2
				    jddd=" 省质监局"
		        case 3
				    jddd=" 市计量局"
				case 4
					jddd=" 计量局"
			end select
			
			dwt.out jddd				 
			dwt.out"</td>"
			dwt.out "      <td class='tdcl tdal'>"
			select case rs("jdzq")
			    case 1
				 	jdzq="6个月"
				case 2
				    jdzq="12个月"
		        case 3
				    jdzq="24个月"
				case 4
					jdzq="48个月"
			end select
			dwt.out jdzq
			dwt.out"</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("bz")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdbr tdal'>"
			call editdel(rs("id"),11,"jlqjtz.asp?action=edit&id=","jlqjtz.asp?action=del&id=")   '1应改为检测科的车间ID
			dwt.out "</td></tr>"
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
	dwt.out " <form method='Get' name='SearchForm' action='jlqjtz.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "  <a href=""jlqjtz.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
	dwt.out "<strong>位号搜索：</strong>" & vbCrLf
	dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	dwt.out "<font color='0066CC'> 查看所属单位的相关内容：</font>"
	dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "<option value=''>按单位跳转至…</option>" & vbCrLf
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