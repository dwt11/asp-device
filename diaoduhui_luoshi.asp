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
dim sqlpxst,rspxst,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel,wangong
url="diaoduhui_luoshi.asp"

keys=trim(request("keyword")) 

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")
if request("wangong")="" then
wangong="全部"
else
wangong=request("wangong")
end if

zr_danwei=request("zr_danwei")
ly=request("ly")




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
  case "del"
    'if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
	call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then
	 call main
end select	

sub add()
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>添 加 任 务</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='diaoduhui_luoshi.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >任务内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_title>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
'/*	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >会议分类:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element >"& vbCrLf
'	dwt.out "<select name='huiyiluoshi_class'>"
'	dwt.out "<option value='0'>请选择分类</option>"& vbCrLf
'
'	dim sql,rs
'	sql="SELECT * from huiyiluoshi_class"
'    set rs=server.createobject("adodb.recordset")
'    rs.open sql,connpxjhzj,1,1
'    if rs.eof then 
'	else
'	do while not rs.eof
'       	response.write"<option value='"&rs("id")&"'>"&rs("class_name")&"</option>"& vbCrLf
'	    'usernameh=rsbz("username1")
'		rs.movenext
'	loop
'	end if 
'	rs.close
'	set rs=nothing
'	dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
'	*/
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>安排时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&session("username1")&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&date()&"'  disabled='disabled' >"& vbCrLf
	dwt.out "<input name='pxst_date' type='hidden' value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >责任单位:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element >"& vbCrLf
'	dwt.out "<select name='zr_danwei'>"
'	dwt.out "<option value='0'>请选择单位</option>"& vbCrLf
'
'	sql="SELECT * from danwei_class"
'    set rs=server.createobject("adodb.recordset")
'    rs.open sql,connpxjhzj,1,1
'    if rs.eof then 
'	else
'	do while not rs.eof
'       	response.write"<option value='"&rs("class_name")&"'>"&rs("class_name")&"</option>"& vbCrLf
'	    'usernameh=rsbz("username1")
'		rs.movenext
'	loop
'	end if 
'	rs.close
'	set rs=nothing
'	dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >责任单位:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	'dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei>"& vbCrLf
	%>
    
    <input name="z1" type="checkbox" value="维修一" /> 维修一 &nbsp;&nbsp;&nbsp;
    <input name="z2" type="checkbox" value="维修二" />维修二&nbsp;&nbsp;&nbsp;
    <input name="z3" type="checkbox" value="维修三 " />维修三&nbsp;&nbsp;&nbsp;
    <input name="z4" type="checkbox" value="维修四" />维修四&nbsp;&nbsp;&nbsp;
    <input name="z5" type="checkbox" value="综合" />综合&nbsp;&nbsp;&nbsp;
    <input name="z6" type="checkbox" value="计量" />计量&nbsp;&nbsp;&nbsp;
    <input name="z7" type="checkbox" value="技术科" />技术科&nbsp;&nbsp;&nbsp;
        <input name="z8" type="checkbox" value="办公室" />办公室&nbsp;&nbsp;&nbsp;
        <input name="z9" type="checkbox" value="检测科" />检测科&nbsp;&nbsp;&nbsp;
    
    <%
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >来源:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	'dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei>"& vbCrLf
	%>
    <select name="ly">
    <option value="周二调度会">周二调度会</option>
    <option value="周五生产例会">周五生产例会</option>
    <option value="党工团">党工团</option>
    <option value="机动例会">机动例会</option>
    <option value="环保例会">环保例会</option>
    <option value="培训例会">培训例会</option>
    <option value="车间汇报">车间汇报</option>
    <option value="安全例会">安全例会</option>
<option value="公司调度会">公司调度会</option>
<option value="设备管理系统重点工作任务书">设备管理系统重点工作任务书</option>
<option value="计量">计量</option>
</select>    
    <%
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >责任人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_ren>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>要求完工时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='yaoqiu_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


					  
'	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
'	    Dwt.out "<iframe src='neweditor/editor.htm?id=pxst_body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
'       
'      dwt.out"  <input type='hidden' name='pxst_body' value=''>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
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
	dwt.out"</div> "& vbCrLf  
	
end sub	

sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from huiyiluoshi" 
      rsadd.open sqladd,conne,1,3
      rsadd.addnew
      
	  
	  zr_danwei= request("z1")&" "&request("z2")&" "&request("z3")&" "&request("z4")&" "&request("z5")&" "&request("z6")&" "&request("z7")&" "&request("z8")&" "&request("z9")
	  'dwt.out zr_danwei
	  rsadd("pxst_title")=Trim(Request("pxst_title"))
      'rsadd("pxst_zz")=request("pxst_zz")
      rsadd("pxst_body")=Trim(request("pxst_body"))
      rsadd("pxst_date")=request("pxst_date")
     ' rsadd("pxst_class")=request("huiyiluoshi_class")
      rsadd("userid")=session("userid")
      rsadd("huiyi_date")=request("txdate")
      rsadd("edit_date")=request("pxst_date")
      rsadd("ly")=request("ly")
      rsadd("zr_danwei")=zr_danwei
      rsadd("yaoqiu_date")=request("yaoqiu_date")
      rsadd("zr_ren")=request("zr_ren")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='diaoduhui_luoshi.asp';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from huiyiluoshi where id="&id
   rsedit.open sqledit,conne,1,1
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>编 辑 任 务</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='diaoduhui_luoshi.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >任务内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_title value='"&rsedit("pxst_title")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

'
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >会议分类:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element >"& vbCrLf
'	dwt.out "<select name='huiyiluoshi_class'>"
'	dwt.out "<option value='0'>请选择分类</option>"& vbCrLf
'
'	dim sql,rs
'	sql="SELECT * from huiyiluoshi_class"
'    set rs=server.createobject("adodb.recordset")
'    rs.open sql,connpxjhzj,1,1
'    if rs.eof then 
'	else
'	do while not rs.eof
'       	response.write"<option value='"&rs("id")&"' "
'		if rsedit("pxst_class")=rs("id") then dwt.out "selected"
'		dwt.out ">"&rs("class_name")&"</option>"& vbCrLf
'	    'usernameh=rsbz("username1")
'		rs.movenext
'	loop
'	end if 
'	rs.close
'	set rs=nothing
'	dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>安排时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("huiyi_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&usernameh(rsedit("userid"))&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&rsedit("pxst_date")&"'  disabled='disabled' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
'	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px' >责任单位:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class=x-form-element >"& vbCrLf
'	dwt.out "<select name='zr_danwei'>"
'	dwt.out "<option value='0'>请选择单位</option>"& vbCrLf
'
'	sql="SELECT * from danwei_class"
'    set rs=server.createobject("adodb.recordset")
'    rs.open sql,connpxjhzj,1,1
'    if rs.eof then 
'	else
'	do while not rs.eof
'       	response.write"<option value='"&rs("class_name")&"' "
'		if rsedit("zr_danwei")=rs("class_name") then dwt.out "selected"
'		dwt.out ">"&rs("class_name")&"</option>"& vbCrLf
'	    'usernameh=rsbz("username1")
'		rs.movenext
'	loop
'	end if 
'	rs.close
'	set rs=nothing
'	dwt.out"</select>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
	
	
	
	
	
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >责任单位:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	'dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei>"& vbCrLf
	%>
    
    <input name="z1" type="checkbox" value="维修一" /> 维修一 &nbsp;&nbsp;&nbsp;
    <input name="z2" type="checkbox" value="维修二" />维修二&nbsp;&nbsp;&nbsp;
    <input name="z3" type="checkbox" value="维修三 " />维修三&nbsp;&nbsp;&nbsp;
    <input name="z4" type="checkbox" value="维修四" />维修四&nbsp;&nbsp;&nbsp;
    <input name="z5" type="checkbox" value="综合" />综合&nbsp;&nbsp;&nbsp;
    <input name="z6" type="checkbox" value="计量" />计量&nbsp;&nbsp;&nbsp;
    <input name="z7" type="checkbox" value="技术科" />技术科&nbsp;&nbsp;&nbsp;
        <input name="z8" type="checkbox" value="办公室" />办公室&nbsp;&nbsp;&nbsp;
    
        <input name="z98" type="checkbox" value="检测科" />检测科&nbsp;&nbsp;&nbsp;
    <%
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >来源:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	'dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei>"& vbCrLf
	%>
    <select name="ly">
    <option value="周二调度会" <%  if rsedit("ly")="周二调度会" then %>selected="selected"<%end if %>>周二调度会</option>
    <option value="周五生产例会" <%  if rsedit("ly")="周五生产例会" then %>selected="selected"<%end if %>>周五生产例会</option>
    <option value="党工团" <%  if rsedit("ly")="党工团" then %>selected="selected"<%end if %>>党工团</option>
    <option value="机动例会" <%  if rsedit("ly")="机动例会" then %>selected="selected"<%end if %>>机动例会</option>
    <option value="环保例会" <%  if rsedit("ly")="环保例会" then %>selected="selected"<%end if %>>环保例会</option>
    <option value="培训例会" <%  if rsedit("ly")="培训例会" then %>selected="selected"<%end if %>>培训例会</option>
    <option value="车间汇报" <%  if rsedit("ly")="车间汇报" then %>selected="selected"<%end if %>>车间汇报</option>
    <option value="安全例会" <%  if rsedit("ly")="安全例会" then %>selected="selected"<%end if %>>安全例会</option>
<option value="公司调度会"  <%  if rsedit("ly")="公司调度会" then %>selected="selected"<%end if %>>公司调度会</option>
<option value="设备管理系统重点工作任务书"  <%  if rsedit("ly")="设备管理系统重点工作任务书" then %>selected="selected"<%end if %>>设备管理系统重点工作任务书</option>
<option value="计量"  <%  if rsedit("ly")="计量" then %>selected="selected"<%end if %>>计量</option>

</select>    







    <%
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	
	
	
	
	

	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >责任人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_ren value='"&rsedit("zr_ren")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>要求完工时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='yaoqiu_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("yaoqiu_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>实际完工时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='wangong_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("wangong_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

				  
'	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
'	dwt.out"				<LABEL style='WIDTH: 75px'>内容:</LABEL>"& vbCrLf
'	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
'	scontent=rsedit("pxst_body")
'	    Dwt.out "<iframe src='neweditor/editor.htm?id=pxst_body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
'   dwt.out "<textarea name='pxst_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
'	dwt.out"				</DIV>"& vbCrLf
'	dwt.out"			  </DIV>"& vbCrLf
'	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >完工情况:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	    dwt.out"<input name='wanggong' type='radio' value='true'"
	if rsedit("isno")=true then dwt.out "checked"
	dwt.out" />已完工 <input name='wanggong' type='radio' value='false' "
	if rsedit("isno")=false then dwt.out "checked"
	dwt.out" />未完工"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 保存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	dwt.out"</div> "& vbCrLf  


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from huiyiluoshi where ID="&ReplaceBadChar(Trim(request("ID")))
	  zr_danwei= trim(request("z1")&" "&request("z2")&" "&request("z3")&" "&request("z4")&" "&request("z5")&" "&request("z6")&" "&request("z7")&" "&request("z8"))
rsedit.open sqledit,conne,1,3
rsedit("pxst_title")=ReplaceBadChar(Trim(Request("pxst_title")))
rsedit("pxst_body")=Trim(request("pxst_body"))
'rsedit("pxst_class")=Trim(request("huiyiluoshi_class"))
rsedit("edit_date")=date()
rsedit("huiyi_date")=request("txdate")
rsedit("yaoqiu_date")=request("yaoqiu_date")
rsedit("wangong_date")=request("wangong_date")
rsedit("ly")=request("ly")
if zr_danwei<>"" then 
  rsedit("zr_danwei")=zr_danwei
'  response.Write zr_danwei
end if 
rsedit("zr_ren")=request("zr_ren")
rsedit("isno")=trim(request("wanggong"))
rsedit.update
rsedit.close
	  dwt.out"<Script Language=Javascript>location.href='diaoduhui_luoshi.asp';</Script>"
end sub


sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>工作落实"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
'	search()
'dim sqlcj,rscj
dwt.out "<div class='x-toolbar'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='diaoduhui_luoshi.asp'>" & vbCrLf
dwt.out "<a href=""diaoduhui_luoshi.asp?action=add"">添加任务</a>&nbsp;&nbsp;标题搜索：" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
		 dwt.out "<select name='wangong'>" 
	dwt.out "<option value='全部'"
	if wangong="全部" then dwt.out "selected"
	dwt.out ">全部</option>" 
	dwt.out "<option value='完工'"
	if wangong="完工" then dwt.out "selected"
	dwt.out ">完工</option>" 
	dwt.out "<option value='未完工'"
	if wangong="未完工" then dwt.out "selected"
	dwt.out ">未完工</option>" 
    dwt.out "</select>" 





		 dwt.out "<select name='zr_danwei'>" 
	dwt.out "<option value=''"
	if zr_danwei="" then dwt.out "selected"
	dwt.out ">责任单位</option>" 
	
	
	dwt.out "<option value='维修一'"
	if zr_danwei="维修一" then dwt.out "selected"
	dwt.out ">维修一</option>" 
	
	
	dwt.out "<option value='维修二'"
	if zr_danwei="维修二" then dwt.out "selected"
	dwt.out ">维修二</option>" 
    
	dwt.out "<option value='维修三'"
	if zr_danwei="维修三" then dwt.out "selected"
	dwt.out ">维修三</option>" 
    
	dwt.out "<option value='维修四'"
	if zr_danwei="维修四" then dwt.out "selected"
	dwt.out ">维修四</option>" 
    
	dwt.out "<option value='综合'"
	if zr_danwei="综合" then dwt.out "selected"
	dwt.out ">综合</option>" 
    
	
	dwt.out "<option value='计量'"
	if zr_danwei="计量" then dwt.out "selected"
	dwt.out ">计量</option>" 
    
	
	dwt.out "<option value='技术科'"
	if zr_danwei="技术科" then dwt.out "selected"
	dwt.out ">技术科</option>" 
    
	
	dwt.out "<option value='办公室'"
	if zr_danwei="办公室" then dwt.out "selected"
	dwt.out ">办公室</option>" 
    
	
	
	
	
	
	dwt.out "<option value='检测科'"
	if zr_danwei="检测科" then dwt.out "selected"
	dwt.out ">检测科</option>" 
    
	
	
	
	
	
	dwt.out "</select>" 

	%>
	
	
	
	    <select name="ly">
    <option value="" <%  if ly="" then %>selected="selected"<%end if %>>来源</option>
    <option value="周二调度会" <%  if ly="周二调度会" then %>selected="selected"<%end if %>>周二调度会</option>
    <option value="周五生产例会" <%  if ly="周五生产例会" then %>selected="selected"<%end if %>>周五生产例会</option>
    <option value="党工团" <%  if ly="党工团" then %>selected="selected"<%end if %>>党工团</option>
    <option value="机动例会" <%  if ly="机动例会" then %>selected="selected"<%end if %>>机动例会</option>
    <option value="环保例会" <%  if ly="环保例会" then %>selected="selected"<%end if %>>环保例会</option>
    <option value="培训例会" <%  if ly="培训例会" then %>selected="selected"<%end if %>>培训例会</option>
    <option value="车间汇报" <%  if ly="车间汇报" then %>selected="selected"<%end if %>>车间汇报</option>
    <option value="安全例会" <%  if ly="安全例会" then %>selected="selected"<%end if %>>安全例会</option>
<option value="公司调度会"  <%  if ly="公司调度会" then %>selected="selected"<%end if %>>公司调度会</option>
<option value="设备管理系统重点工作任务书"  <%  if ly="设备管理系统重点工作任务书" then %>selected="selected"<%end if %>>设备管理系统重点工作任务书</option>
<option value="计量"  <%  if ly="计量" then %>selected="selected"<%end if %>>计量</option>

</select>    

	
	<%
	



dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
dwt.out "</form></div>" & vbCrLf

'	sqlpxst="SELECT * from huiyiluoshi " 
'	if request("classid")<>"" then sqlpxst=sqlpxst&" where pxst_class="&request("classid")&" and pxst_body like '%"&keys&"%'"
'	sqlpxst=sqlpxst&" ORDER BY id DESC"
	
		sqlpxst="SELECT * from huiyiluoshi where 1=1 "

	
	

	if wangong="完工" then 
	sqlpxst=sqlpxst& " and  isno=true  "
	end if
	if wangong="未完工" then 
	sqlpxst=sqlpxst& " and  isno=false  "
	end if

	if keys<>"" then 
	sqlpxst=sqlpxst& "  and  pxst_body like '%" &keys& "%'  "
	end if 
	
	if zr_danwei<>"" then 
	sqlpxst=sqlpxst& "  and  zr_danwei like '%" &zr_danwei& "%'  "
	end if 
	if ly<>"" then 
	sqlpxst=sqlpxst& " and  ly='"&ly&"' "
	end if
	
		sqlpxst=sqlpxst&"  ORDER BY id DESC"
'dwt.out sqlpxst
	set rspxst=server.createobject("adodb.recordset")
	rspxst.open sqlpxst,conne,1,1
	if rspxst.eof and rspxst.bof then 
	message("未找到相关任务")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td' ><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>来源</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>安排时间</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>要求时间</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>完工时间</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>任  务  标  题</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>责任单位</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>责任人</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布时间</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>完成情况</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
		dwt.out "    </tr>"
           record=rspxst.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxst.PageSize = Cint(PgSz) 
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
           rspxst.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxst.PageSize
           do while not rspxst.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("ly")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("huiyi_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("yaoqiu_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("wangong_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"
				 
				 
				 
				 
				 dwt.out"<a href=diaoduhui_view.asp?id="&rspxst("id")&" target=_blank>"
				 
				 	if rspxst("isno")=true then

				dwt.out rspxst("pxst_title")
				else
				dwt.out "<strong><div style=' color: #F00;'>"&rspxst("pxst_title")&"</div></strong>"
				end if  
				 
				
				 
				 dwt.out"</a></td>"
         dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("zr_danwei")&"&nbsp;</td>"
          dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("zr_ren")&"&nbsp;</td>"
           
				 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 if isnull(rspxst("userid")) then 
				   dwt.out rspxst("pxst_zz")
				 else
				   dwt.out usernameh(rspxst("userid")) 
				 end if   
				 dwt.out"&nbsp;</div></td>"
				 
                 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_date")&"</div></td>"
				 
	if rspxst("isno")=true then
		dwt.out "<td  style=""border-bottom-style: solid;border-width:1px;white-space:nowrap"" ><div align=""center"">完成</div></td>"
	else
		dwt.out "<td  style=""border-bottom-style: solid;border-width:1px;white-space:nowrap"" ><div align=""center"">未完成</div></td>"
	end if
				 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 
				 			if session("levelclass")<>0 then 
				dim rsqxtb_fk,sqlqxtb_fk
				set rsqxtb_fk=server.createobject("adodb.recordset")
				sqlqxtb_fk="select * from huiyiluoshi_fk where huiyiluoshi_fk_sscj="&session("levelclass")&" and huiyiluoshi_id="&rspxst("id")
				rsqxtb_fk.open sqlqxtb_fk,conne,1,1
				if rsqxtb_fk.eof and rsqxtb_fk.bof then 
					dwt.out  "<a href='huiyiluoshi_fk.asp?action=add&huiyiluoshi_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"'>添加反馈</a>&nbsp;"
				else
					dwt.out  "<a href='huiyiluoshi_fk.asp?action=edit&huiyiluoshi_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"'>编辑反馈</a>&nbsp;"
					if session("level")=0 then dwt.out  "<a href='huiyiluoshi_fk.asp?action=del&qxtb_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"' onClick=""return confirm('确定要删除此反馈吗？');"">删除反馈</a>"
				end if 
				rsqxtb_fk.close
				set rsqxtb_fk=nothing
			end if 

				 
				 
				 if session("level")=0 or session("levelclass")=9 or rspxst("userid")=session("userid") then
				  dwt.out "<a href='diaoduhui_luoshi.asp?action=edit&ID="&rspxst("id")&"'>编辑</a>"
				  dwt.out "&nbsp;<a href='diaoduhui_luoshi.asp?action=del&ID="&rspxst("id")&"' onClick=""return confirm('确定要删除此试题吗？');"">删除</a>"
				 end if 			'call editdel(rspxst("id"),rspxst("sscj"),"diaoduhui_luoshi.asp?action=edit&id=","diaoduhui_luoshi.asp?action=del&id=")
				 dwt.out "&nbsp; </div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rspxst.movenext
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
	rspxst.close
	set rspxst=nothing
	conn.close
	set conn=nothing
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from huiyiluoshi where id="&id
rsdel.open sqldel,conne,1,3
dwt.out"<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub


Function class_name(class_id)
    dim sqlcj,rscj
'dim class_id

	  sqlcj="SELECT * from huiyiluoshi_class where id="&class_id
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,connpxjhzj,1,1
    if rscj.eof then 
		class_name="未编辑"
	else
	do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    class_name=rscj("class_name")
		rscj.movenext
	loop
	end if 
	rscj.close
	set rscj=nothing
end Function
dwt.out "</body></html>"


sub search()
dim sqlcj,rscj
dwt.out "<div class='x-toolbar'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='diaoduhui_luoshi.asp'>" & vbCrLf
dwt.out "<a href=""diaoduhui_luoshi.asp?action=add"">添加任务</a>&nbsp;&nbsp;标题搜索：" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
dwt.out "</form></div>" & vbCrLf
end sub

Call CloseConn
%>