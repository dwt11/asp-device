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
dim sqlnews,rsnews,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
dim sscjid,rssscj,sqlsscj
classid=request("classid")
if classid="" then classid=50
    url=geturl
    classname=conna.Execute("SELECT class_name FROM xzgl_news_class WHERE id="&classid)(0)
keys=trim(request("keyword"))
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统--安全生产责任制</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/tab.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
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
  	dwt.out"<DIV style='WIDTH: 780px;padding-top:50px;padding-left:50px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<div align=center><H3 style='MARGIN-BOTTOM: 5px'>添加"&classname&"</H3></div>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' action='aqgl.asp' name='form1'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>是否记录查看人员:</DIV></div></LABEL>"& vbCrLf
	dwt.out"				<DIV >"& vbCrLf
	dwt.out"				  <div align=left><INPUT name=isviewd type=checkbox></DIV>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>标题:</DIV></div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_title>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>分类:</DIV></LABEL>"& vbCrLf
		dwt.out"				<DIV class=x-form-element align=left>"& vbCrLf
		dwt.out "<select name='news_class'>"
		dwt.out "<option value='0'>请选择分类</option>"& vbCrLf
		dim sql,rs
		sql="SELECT * from xzgl_news_class"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connxzgl,1,1
		if rs.eof then 
		else
		do while not rs.eof
			dwt.out "<option value='"&rs("id")&"'"
			if cint(request("classid"))=rs("id") then dwt.out " selected"
			dwt.out ">"&rs("class_name")&"</option>"& vbCrLf
			rs.movenext
		loop
		end if 
		rs.close
		set rs=nothing
		dwt.out"</select>"
		dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
					  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>发布人:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_title  disabled='disabled' value="&session("username1")&">"& vbCrLf
	dwt.out"<input name='user_id' type='hidden' value="&session("userid")&"></td></tr>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>车间:</DIV></LABEL>"& vbCrLf
		dwt.out"				<DIV class=x-form-element align=left>"& vbCrLf
		dwt.out "<select name='sscj'>"
		'dwt.out "<option value='0'>请选择车间</option>"& vbCrLf
		sqlcj="SELECT * from levelname where levelclass<>10 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
		dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
		rscj.movenext
		loop
		rscj.close
		set rscj=nothing		
		dwt.out"</select>"
		dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>时间:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element  style='PADDING-LEFT: 80px'>"& vbCrLf
    'dwt.out"<input name='news_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
dwt.out"<input name=news_date style='WIDTH: 175px'  disabled='disabled'  value='"&NOW()&"'>"	
dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"				<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>内容:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=news_body&style=s_blue&originalfilename=d_originalfilename &savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='550' HEIGHT='350'>"
	 dwt.out "</iframe>  <input type='hidden' name='news_body' value=''>"	

	
	 dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'><div align=center>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>      <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </div></DIV>"& vbCrLf
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
	
   
   
end sub	






sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xzgl_news" 
      rsadd.open sqladd,connxzgl,1,3
      rsadd.addnew
      rsadd("news_title")=ReplaceBadChar(Trim(Request("news_title")))
      rsadd("user_id")=request("user_id")
      rsadd("sscj")=Trim(request("sscj"))
      rsadd("news_body")=Trim(request("news_body"))
      rsadd("news_date")=NOW()
      rsadd("news_class")=request("news_class")
      if request("isviewd")="on" then rsadd("isviewd")=true
      rsadd.update
      rsadd.close
	dwt.savesl conna.Execute("SELECT class_name FROM xzgl_news_class WHERE id="&request("news_class"))(0) ,"添加",ReplaceBadChar(Trim(Request("news_title")))
      set rsadd=nothing
	  dwt.out "<Script Language=Javascript>location.href='aqgl.asp?classid="&request("news_class")&"';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from xzgl_news where id="&id
   rsedit.open sqledit,connxzgl,1,1
  	dwt.out"<DIV style='WIDTH: 780px;padding-top:50px;padding-left:50px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<div align=center><H3 style='MARGIN-BOTTOM: 5px'>编辑"&classname&"</H3></div>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' action='aqgl.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>是否记录查看人员:</DIV></div></LABEL>"& vbCrLf
	dwt.out"				<DIV >"& vbCrLf
	dwt.out"				  <div align=left><INPUT name=isviewd type=checkbox "
	if rsedit("isviewd") then dwt.out "checked"
	dwt.out"></DIV>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>标题:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_title value="&rsedit("news_title")&">"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>分类:</div></LABEL>"& vbCrLf
		dwt.out"				<DIV class=x-form-element >"& vbCrLf
		dwt.out "<select name='news_class'>"
		dwt.out "<option value='0'>请选择分类</option>"& vbCrLf
		dim sql,rs
		sql="SELECT * from xzgl_news_class"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connxzgl,1,1
		if rs.eof then 
		else
		do while not rs.eof
			dwt.out "<option value='"&rs("id")&"'"
			if cint(rsedit("news_class"))=rs("id") then dwt.out " selected"
			dwt.out ">"&rs("class_name")&"</option>"& vbCrLf
			'usernameh=rsbz("username1")
			rs.movenext
		loop
		end if 
		rs.close
		set rs=nothing
		dwt.out"</select>"
		dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>发布人:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_title  disabled='disabled' value="&usernameh(rsedit("user_id"))&">"& vbCrLf
	dwt.out"<input name='user_id' type='hidden' value="&rsedit("user_id")&"></td></tr>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>所属车间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
		dwt.out"<select name='sscj' style='WIDTH: 175px' size='1'>"& vbCrLf
		dwt.out"<option  selected value='"&rsedit("sscj")&"'>"&sscjh(rsedit("sscj"))&"</option>"& vbCrLf
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
		dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		
		 	 
	
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    'dwt.out"<input name='news_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("news_date")&"'>"
dwt.out"<input name='news_date' style='WIDTH: 175px'  disabled='disabled'  value='"&NOW()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out""& vbCrLf '此处有问题
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>内容:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	scontent=rsedit("news_body")
dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=news_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='550' HEIGHT='350'>"

   dwt.out "</iframe><textarea name='news_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
	 dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'><div align=center>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'>  	<input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </div></DIV>"& vbCrLf
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


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from xzgl_news where ID="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connxzgl,1,3
	rsedit("news_title")=ReplaceBadChar(Trim(Request("news_title")))
	rsedit("news_body")=Trim(request("news_body"))
	rsedit("sscj")=Trim(Request("sscj"))
	rsedit("news_date")=NOW()
	rsedit("news_class")=request("news_class")
      if request("isviewd")="on" then rsedit("isviewd")=true
	rsedit.update
	rsedit.close
	dwt.savesl conna.Execute("SELECT class_name FROM xzgl_news_class WHERE id="&request("news_class"))(0) ,"编辑",ReplaceBadChar(Trim(Request("news_title")))
	dwt.out "<Script Language=Javascript>history.go(-2);</Script>"
	
end sub


sub main()
		sscjid=request("sscj")
		if sscjid="" then sscjid=1
		sqlbody="SELECT * from xzgl_news where news_class="&classid
	if keys<>"" then 
		sqlbody=sqlbody&" and name like '%" &keys& "%' "
		title=title&"-搜索 "&keys
	end if
	if sscjid<>"" then
		sqlbody=sqlbody&" and sscj="&sscjid
		title=title&"-"&sscjh(sscjid)
	end if 
	sqlbody=sqlbody&" order by ID desc"
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>安全生产责任制"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	call search()

	dwt.out "<div class='navg'>"
	dwt.out "  <div id='system' class='mainNavg'>"
	dwt.out "    <ul>"
	sqlsscj="SELECT * from levelname where levelclass<>0"
	set rssscj=server.createobject("adodb.recordset")
	rssscj.open sqlsscj,conn,1,1
	if rssscj.eof and rssscj.bof then 
		dwt.out  message ("<p align='center'>未添加生产车间</p>" )
	else
	do while not rssscj.eof 
		if cint(sscjid)=rssscj("levelid") then 
		   dwt.out "<li id='systemNavg'><a href='#'>"&rssscj("levelname")&"</a></li>"
		else
		   dwt.out "<li><a href='aqgl.asp?sscj="&rssscj("levelid")&"'>"&rssscj("levelname")&"</a></li>"
		end if    
	rssscj.movenext
	loop
	end if 
	rssscj.close
	set rssscj=nothing
	  
	  
    dwt.out "</ul>"
    dwt.out " </div>"
	
	
	set rsnews=server.createobject("adodb.recordset")
	rsnews.open sqlbody,conna,1,1
	if rsnews.eof and rsnews.bof then 
		dwt.out  "<div class='textbody1'>"
		dwt.out  message ("<p align='center'>未添加内容</p>" )
		dwt.out  "</div>"
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out  "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>标题</div></td>"
		dwt.out  "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>发布时间</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
		dwt.out  "    </tr>"
	   record=rsnews.recordcount
	   if Trim(Request("PgSz"))="" then
		   PgSz=20
	   ELSE 
		   PgSz=Trim(Request("PgSz"))
	   end if 
	   rsnews.PageSize = Cint(PgSz) 
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
	   rsnews.absolutePage = page
	   start=PgSz*Page-PgSz+1
	   rowCount = rsnews.PageSize
	   do while not rsnews.eof and rowcount>0
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			 dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh_id&"</div></td>"
			 dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</a></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			  if rsnews("user_id")=0 then 
				dwt.out rsnews("news_zz")
			  else
				dwt.out usernameh(rsnews("user_id"))
			  end if
			 
			 dwt.out "</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsnews("news_date")&"</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sscjh(rsnews("sscj"))&"</div></td>"
			 dwt.out  "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			 if rsnews("user_id")=0 or session("userid")=79 or session("userid")=1 or rsnews("user_id")=session("userid") then 
				 dwt.out  "<a href='aqgl.asp?classid="&classid&"&action=edit&ID="&rsnews("id")&"'>编辑</a>&nbsp;"
				 dwt.out "<a href='aqgl.asp?action=del&ID="&rsnews("id")&"' onClick=""return confirm('确定要删除此内容吗？');"">删除</a>"
			 end if 
			 dwt.out  "&nbsp; </div></td>"
			 dwt.out  "    </tr>"
			 RowCount=RowCount-1
	  rsnews.movenext
	  loop
	dwt.out  "</table>"
   call showpage(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
   rsnews.close
   set rsnews=nothing
	conn.close
	set conn=nothing
end sub



sub del()
ID=request("ID")


	sqledit="select * from xzgl_news where ID="&id

	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from xzgl_news where id="&id
	rsdel.open sqldel,connxzgl,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	'rsdel.close
	set rsdel=nothing  

end sub

sub search()
	dim sqlcj,rscj
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='aqgl.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.out "<a href=""aqgl.asp?action=add&classid="&classid&""">添加</a>"
	dwt.out "&nbsp;&nbsp;搜索：" & vbCrLf
	Dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	Dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf
end sub







dwt.out  "</body></html>"

Call CloseConn
%>