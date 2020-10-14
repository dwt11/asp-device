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
url=geturl
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统检修记录管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.jxjl_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form1.jxjl_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.jxjl_body.value==''){" & vbCrLf
dwt.out "      alert('检修内容不能为空！');" & vbCrLf
dwt.out "   document.form1.jxjl_body.focus();" & vbCrLf
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


sub add()
dim sqlcj,rscj

	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>添加检修记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='jxjl.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	if session("level")=0 then 
		dwt.out"<select name='jxjl_sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
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
		dwt.out"<input name='jxjl_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
	end if 
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >检修原因:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jxjl_jxyy >请添写检修原因</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >检修内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jxjl_body >请添写检修内容</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"							<DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检修人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检修时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jxjl_jxdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>备注:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_bz>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
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
	dwt.out"</DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  
	
   
   
end sub	

sub saveadd()  
dim sqladd,rsadd  
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from jxjl" 
      rsadd.open sqladd,conndcs,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("jxjl_sscj"))
      rsadd("jxyy")=request("jxjl_jxyy")
      rsadd("body")=Trim(request("jxjl_body"))
      rsadd("jxrname")=request("jxjl_jxrname")
      rsadd("jxdate")=request("jxjl_jxdate")
      rsadd("bz")=request("jxjl_bz")
      rsadd("userid")=session("userid")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub saveedit()  
dim rsedit,sqledit  
	  '保存
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jxjl where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conndcs,1,3
      'rsedit("sscj")=Trim(Request("jxjl_sscj"))
      rsedit("jxyy")=request("jxjl_jxyy")
      rsedit("body")=Trim(request("jxjl_body"))
      rsedit("jxrname")=request("jxjl_jxrname")
      rsedit("jxdate")=request("jxjl_jxdate")
      rsedit("bz")=request("jxjl_bz")
      rsedit("userid")=session("userid")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from jxjl where id="&id
  rsdel.open sqldel,conndcs,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub

sub edit()
  	 

   
   dim sqledit,rsedit,id
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from jxjl where id="&id
   rsedit.open sqledit,conndcs,1,1
   dwt.out"<br><form method='post' action='jxjl.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑检修记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
   dwt.out"<form method='post' action='jxjl.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	
	dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px'  value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >检修原因:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jxjl_jxyy >"&rsedit("jxyy")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >检修内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=jxjl_body >"&rsedit("body")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"							<DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检修人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>检修时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jxjl_jxdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("jxdate")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>备注:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_bz value="&rsedit("bz")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	dim sqljxjl,rsjxjl,title
	sqljxjl="SELECT * from jxjl"
	if keys<>"" then 
		sqljxjl=sqljxjl&" where body like '%" &keys& "%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
		sqljxjl=sqljxjl&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	sqljxjl=sqljxjl&" ORDER BY sscj aSC,jxdate desc"
	
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>检修记录"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	
	'for sscji=1 to 5 '071017修改
	sql="select * from levelname where istq=false"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		dwt.out "没有添加车间"
	else
	   do while not rs.eof
		sql="SELECT count(id) FROM jxjl WHERE sscj="&rs("levelid")&" and month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
		numb=numb&sscjh_d(rs("levelid"))&":"&"<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
	rs.movenext
	loop
	end if 
	rs.close
	
	sql="SELECT count(id) FROM jxjl WHERE  month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
	totall= "<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>" 
	dwt.out "<div class='pre'>本月"&numb&"合计:"&totall&"</div>"& vbCrLf

	search()
	
	set rsjxjl=server.createobject("adodb.recordset")
	rsjxjl.open sqljxjl,conndcs,1,1
	if rsjxjl.eof and rsjxjl.bof then 
		message("未找到相关检修记录")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>检修原因</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>检修内容</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>检修人</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>时间</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>备注</div></td>"& vbCrLf
		'dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"& vbCrLf
		dwt.out "    </tr>"& vbCrLf
		record=rsjxjl.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rsjxjl.PageSize = Cint(PgSz) 
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
		rsjxjl.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rsjxjl.PageSize
		do while not rsjxjl.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_d(rsjxjl("sscj"))&"</div></td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjxjl("jxyy")&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&searchH(uCase(rsjxjl("body")),keys)&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjxjl("jxrname")&"&nbsp;</div></td>"& vbCrLf
			dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rsjxjl("jxdate")&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjxjl("bz")&"&nbsp;</td>"& vbCrLf
			'dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"& vbCrLf
			'call editdel(rsjxjl("id"),rsjxjl("sscj"),"jxjl.asp?action=edit&id=","jxjl.asp?action=del&id=")
			'dwt.out "</div></td>"
DWT.OUT "</tr>"& vbCrLf
			RowCount=RowCount-1
			rsjxjl.movenext
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
	rsjxjl.close
	set rsjxjl=nothing
	conn.close
	set conn=nothing
end sub

dwt.out "</body></html>"

sub search()
dim sqlcj,rscj
dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='jxjl.asp'>" & vbCrLf

'if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "<a href=""jxjl.asp?action=add"">添加检修记录</a>&nbsp;&nbsp;"

dwt.out "内容搜索：" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
dwt.out "查看所属车间的相关内容："
dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
        dwt.out"<option value='jxjl.asp?sscj="&rscj("levelid")&"'"
		if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
		dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	dwt.out "</select>&nbsp;&nbsp;<a href=tocsv.asp?action=dcsjxmain&sql1=jxjl&titlename=检修记录>输出检修记录到Excel文档</a>	" & vbCrLf
dwt.out "</form></div></div>" & vbCrLf
end sub





Call CloseConn
%>