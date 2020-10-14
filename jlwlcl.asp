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
dim id,scontent,rsdel,sqldel,wlcl_m,wlcl_d,wlcl_y   

url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>产品产量台帐</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.wlcl_1.value==''){" & vbCrLf
dwt.out "      alert('请输入产量！');" & vbCrLf
dwt.out "   document.form1.wlcl_1.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.wlcl_2.value==''){" & vbCrLf
dwt.out "      alert('请输入产量！');" & vbCrLf
dwt.out "   document.form1.wlcl_1.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.wlcl_3.value==''){" & vbCrLf
dwt.out "      alert('请输入产量！');" & vbCrLf
dwt.out "   document.form1.wlcl_1.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.wldw.value==''){" & vbCrLf
dwt.out "      alert('请输入单位！');" & vbCrLf
dwt.out "   document.form1.wldw.focus();" & vbCrLf
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
	dwt.out "<form method='post' class='x-form' action='jlwlcl.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	
    dwt.out outdatadict ("jdzq","物料名称",onnumb)
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>单位:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wldw' type='text'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(0:00-8:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_1' type='text' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(8:00-16:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_2' type='text' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(16:00-24:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_3' type='text' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>日期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='wldate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)'  >"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf	
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
					  	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>填报人:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name='tbry'  disabled='disabled' value="&session("username1")&">"& vbCrLf
	dwt.out"<input name='user_id' type='hidden' value="&session("userid")&"></td></tr>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>备注:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=wlbz></TEXTAREA>"& vbCrLf
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
	sqladd="select * from wlcl "
	rsadd.open sqladd,connzj,1,3
	wlcl_d=cdbl(trim(request("wlcl_1")))+cdbl(trim(request("wlcl_2")))+cdbl(trim(request("wlcl_3")))
	rsadd.addnew
	on error resume next
	rsadd("jdzq")=request("jdzq")
	rsadd("wlcl_1")=request("wlcl_1")
	rsadd("wlcl_2")=request("wlcl_2")
    rsadd("wlcl_3")=request("wlcl_3")
	
	rsadd("wlcl_d")=wlcl_d
	rsadd("wlcl_m")=wlcl_d
	rsadd("wlcl_y")=wlcl_d

	rsadd("wldw")=request("wldw")
	rsadd("wldate")=request("wldate")
	rsadd("wlbz")=request("wlbz")
    rsadd("user_id")=request("user_id")
    rsadd("sscj")=session("levelclass")
	rsadd.update
	rsadd.close
	Dwt.savesl "产品产量台账","编辑",request("jdzq")
	set rsadd=nothing
	dwt.out"<Script Language=Javascript>location.href='jlwlcl.asp';</Script>"
end sub


sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from wlcl where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connzj,1,3
	wlcl_d=cdbl(trim(request("wlcl_1")))+cdbl(trim(request("wlcl_2")))+cdbl(trim(request("wlcl_3")))
	rsedit("wlcl_1")=request("wlcl_1")
	rsedit("wlcl_2")=request("wlcl_2")
	rsedit("wlcl_3")=request("wlcl_3")
	rsedit("wlcl_d")=wlcl_d
	rsedit("wlcl_m")=wlcl_d
	rsedit("wlcl_y")=wlcl_d
	rsedit("wldw")=request("wldw")
	rsedit("wldate")=request("wldate")
	rsedit("user_id")=request("user_id")
	rsedit("wlbz")=request("wlbz")
	rsedit.update
	rsedit.close
	Dwt.savesl "产品产量台账","编辑",request("jdzq")
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
	id=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from wlcl where id="&id
	rsdel.open sqldel,connzj,1,3
	Dwt.savesl "产品产量台账","删除",id
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub


sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from wlcl where id="&id
	rsedit.open sqledit,connzj,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 370px;padding-top:50px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑产品产量台账</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='jlwlcl.asp' name='form1' onsubmit='javascript:return checkadd();'>"

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>名称:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out outdatadict2 ("jdzq","物料名称",onnumb,rsedit("jdzq"))
'	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='jdzq' type='text' value="&rsedit("jdzq")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>单位:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wldw' type='text' value="&rsedit("wldw")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(0:00-8:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_1' type='text' value="&rsedit("wlcl_1")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(8:00-16:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_2' type='text' value="&rsedit("wlcl_2")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>产量(16:00-24:00):</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <input class='x-form-text x-form-field' style='WIDTH: 175px' name='wlcl_3' type='text' value="&rsedit("wlcl_3")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px'><div align=right>日期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
    dwt.out"<input name='wldate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)'  >"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf	
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
					  	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>填报人:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name='tbry'  disabled='disabled' value="&usernameh(rsedit("user_id"))&">"& vbCrLf
	dwt.out"<input name='user_id' type='hidden' value="&rsedit("user_id")&"></td></tr>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 115px' ><div align=right>备注:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 20px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=wlbz>"&rsedit("wlbz")&"</TEXTAREA>"& vbCrLf
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

	getyear=request("year")
	getmonth=request("month")
	getday=request("day")

    getnowday=date()-1
	
	if getyear="" then getyear=year(getnowday)
	if getmonth="" then getmonth=month(getnowday)
	if getday="" then getday=day(getnowday)



	selectdate=getyear&"-"&getmonth&"-"&getday
	selectdate=cdate(selectdate)
	'message selectdate
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><b>"&selectdate&" 产品产量</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
	dwt.out "		 <form method='post'  action='jlwlcl.asp'  name='form' >"
    	dwt.out "		 <a href='/jlwlcl.asp?action=add'>添加日志</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	dwt.out "<a href='/jlwlcl.asp?year="&year(selectdate-2)&"&month="&month(selectdate-2)&"&day="&day(selectdate-2)&"'>"&year(selectdate-2)&"年"&month(selectdate-2)&"月"&day(selectdate-2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<a href='/jlwlcl.asp?year="&year(selectdate-1)&"&month="&month(selectdate-1)&"&day="&day(selectdate-1)&"'>"&year(selectdate-1)&"年"&month(selectdate-1)&"月"&day(selectdate-1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	dwt.out "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >		 <select name='year'></select>年<select name='month'></select>月<select name='day'></select>日 &nbsp;&nbsp;<input  type='submit' name='Submit' value=' 查看 ' style='cursor:hand;'>"
	dwt.out "		 <script type='text/javascript' src='js/selectdate.js'></script>"
	if now()-selectdate>1 then 	dwt.out "<a href='/jlwlcl.asp?year="&year(selectdate+1)&"&month="&month(selectdate+1)&"&day="&day(selectdate+1)&"'>"&year(selectdate-1)&"年"&month(selectdate+1)&"月"&day(selectdate+1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-selectdate>2 then 	dwt.out "<a href='/jlwlcl.asp?year="&year(selectdate+2)&"&month="&month(selectdate+2)&"&day="&day(selectdate+2)&"'>"&year(selectdate+2)&"年"&month(selectdate+2)&"月"&day(selectdate+2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	
	dwt.out "	</form></div>"
	dwt.out "</div>"

	sql="SELECT * from wlcl where year(wldate)="&getyear&" and month(wldate)="&getmonth&" and day(wldate)="&getday
	if keys<>"" then 
		sql=sql&" and name like '%"&keys&"%' "
		title="-搜索 "&keys
	end if 
	sql=sql&" ORDER BY wldate aSC"

	dwt.out "<div style='left:6px;'>"& vbCrLf
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
	   message "未找到相关内容"
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr bgcolor=""#e5e5e5"">" 
		dwt.out "     <td  class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >序号</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >物料名称</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >单位</div></td>"
		dwt.out "      <td class='x-td' colspan='6'><DIV class='x-grid-hd-text' >产量</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >日期</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >填报人</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >备注</div></td>"
		dwt.out "      <td class='x-td' rowspan='2'><DIV class='x-grid-hd-text' >选项</div></td>"
		dwt.out "    </tr>"

		dwt.out "<tr bgcolor=""#e5e5e5"">" 
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>白班</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>中班</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>夜班</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>日产量</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>月产量</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>年产量</div></td>"
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
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"
					dwt.out dispalydatadict("物料名称",rs("jdzq"))
					dwt.out"&nbsp;</Div></td>" & vbCrLf
'			dwt.out "      <td class='tdcl'>"&rs("jdzq")&"</td>"
			dwt.out "      <td class='tdcl tdbr tdal' >"&rs("wldw")&"</td>"
			dwt.out "      <td class='tdcl'>"&rs("wlcl_1")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl'>"&rs("wlcl_2")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl'>"&rs("wlcl_3")&"&nbsp;</td>"
			wlcl_m=Connzj.Execute("SELECT sum(wlcl_d) FROM wlcl WHERE month(wldate)="&getmonth&" and jdzq="&rs("jdzq"))(0)
			wlcl_y=Connzj.Execute("SELECT sum(wlcl_d) FROM wlcl WHERE year(wldate)="&getyear&" and jdzq="&rs("jdzq"))(0)
			
			dwt.out "      <td class='tdcl'>"&rs("wlcl_d")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl'>"&wlcl_m&"&nbsp;</td>"
			dwt.out "      <td class='tdcl'>"&wlcl_y&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("wldate")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&usernameh(rs("user_id"))&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdal'>"&rs("wlbz")&"&nbsp;</td>"
			dwt.out "      <td class='tdcl tdbr tdal'>"
			call editdel(rs("id"),6,"jlwlcl.asp?action=edit&id=","jlwlcl.asp?action=del&id=")   
			dwt.out "</td></tr>"
			 RowCount=RowCount-1
          rs.movenext
		loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" then
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


Call CloseConn
%>