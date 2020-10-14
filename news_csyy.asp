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



<%
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>信息管理系统畅所欲言</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf


dim sqlnews,rsnews,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel,vip,classid
vip=request.servervariables("http_x_forwarded")
if vip="" then
vip=request.servervariables("remote_addr")
classid=request("classid")
end if
keys=trim(request("keyword")) 
    url=geturl
    classname="畅所欲言"
action=request("action")

select case action
  case "add"
        call add
  case "saveadd"
    call saveadd
  case "edit"
	if session("groupid")=3 then call edit
  case "saveedit"
    call saveedit
  case "del"
    if session("groupid")=3 then call del
  case ""
	 call main
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
	dwt.out "<form method='post' action='news_csyy.asp' name='form1'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf

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
		sql="SELECT * from csyy_class where id=1 "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conncsyy,1,1
		if rs.eof then 
		else
		do while not rs.eof
			dwt.out "<option value='"&rs("id")&"'"
			if cint(request("classid"))=rs("id") then dwt.out " selected"
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
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>发布人:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_zz >"& vbCrLf
	dwt.out"           <input name='user_id' type='hidden' value="&vip&"></td></tr>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>时间:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    'dwt.out"<input name='news_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
dwt.out"<input name='news_date' style='WIDTH: 175px'  disabled='disabled'  value='"&NOW()&"'>"	
dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>内容:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=news_body&style=s_blue&originalfilename=d_originalfilename &savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='550' HEIGHT='350'>"
	 dwt.out "</iframe>  <input type='hidden' name='news_body' value=''>"	
	'DWT.OUT "<input type='hidden' name='news_body' id='news_body'>"& vbCrLf
    'dwt.out "<iframe src='neweditor/editor.htm?id=news_body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"

	
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
     

 if Request("news_title")<>"" then
   	dwt.savesl conncs.Execute("SELECT class_name FROM csyy_class WHERE id="&request("news_class"))(0) ,"添加",ReplaceBadChar(Trim(Request("news_title"))) 

  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from csyy_body" 
      rsadd.open sqladd,conncsyy,1,3
      rsadd.addnew
      rsadd("news_title")=ReplaceBadChar(Trim(Request("news_title")))
      rsadd("user_id")=request("user_id")
	rsadd("news_zz")=Trim(request("news_zz"))
      rsadd("news_body")=Trim(request("news_body"))
      rsadd("news_date")=NOW()
      rsadd("news_class")=request("news_class")
      rsadd.update
      rsadd.close

      set rsadd=nothing
      end if
	  dwt.out "<Script Language=Javascript>location.href='news_csyy.asp?classid="&request("news_class")&"';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from csyy_body where id="&id
   rsedit.open sqledit,conncsyy,1,1
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
	dwt.out "<form method='post' action='news_csyy.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
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
		sql="SELECT * from csyy_class where id=1 "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conncsyy,1,1
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
	if rsedit("news_zz")<>"" then 
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_zz  value="&rsedit("news_zz")&">"& vbCrLf
	else
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=user_id  disabled='disabled' value="&rsedit("user_id")&">"& vbCrLf
	end if
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
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
		'DWT.OUT "<input type='hidden' name='news_body' id='news_body' value='"&scontent&"'>"& vbCrLf
dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=news_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='550' HEIGHT='350'>"
	
	
    'dwt.out "<iframe src='neweditor/editor.htm?id=news_body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"


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
	sqledit="select * from csyy_body where ID="&ReplaceBadChar(Trim(request("ID")))
	
	rsedit.open sqledit,conncsyy,1,3
	rsedit("news_title")=ReplaceBadChar(Trim(Request("news_title")))
	rsedit("user_id")=request("user_id") 
	rsedit("news_zz")=request("news_zz") 
	rsedit("news_body")=Trim(request("news_body"))
	rsedit("news_date")=NOW()
	rsedit("news_class")=request("news_class")
      if request("isviewd")="on" then rsedit("isviewd")=true
	rsedit.update
	rsedit.close
	dwt.savesl conncs.Execute("SELECT class_name FROM csyy_class WHERE id="&request("news_class"))(0) ,"编辑",ReplaceBadChar(Trim(Request("news_title")))
	dwt.out "<Script Language=Javascript>history.go(-2);</Script>"
	
end sub


sub main()
dim sqlcsyy,rscsyy,csyy_numb
    'dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
   ' if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "<a href='news.asp?classid="&classid&"&action=add'>添加"&classname&"</a>&nbsp;&nbsp;"
   ' dwt.out "</div></div>"

	'sqlnews="SELECT * from xzgl_news where news_class="&classid&" ORDER BY id DESC"
	
		'sqlbody="SELECT * from body"
	if keys<>"" then 
		sqlbody="SELECT * from csyy_body  where news_title like '%" &keys& "%'  and news_class="&classid
		title="-搜索 "&keys
	else
	    	sqlbody="SELECT * from csyy_body WHERE news_class="&classid
	
	end if 
	sqlbody=sqlbody&" order by ID desc"

    
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&classname&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	sql="SELECT isre FROM csyy_class WHERE id="&classid
	isre=conncs.Execute(sql)(0)
		
		search()
	set rsnews=server.createobject("adodb.recordset")
	rsnews.open sqlbody,conncs,1,1
	if rsnews.eof and rsnews.bof then 
		dwt.out  message ("<p align='center'>未添加内容</p>" )
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out  "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>标题</div></td>"
		dwt.out  "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out  "      <td  class='x-td'><DIV class='x-grid-hd-text'>浏览"
        if isre then dwt.out "/回复"
		dwt.out "		</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>发布时间</div></td>"
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
			 dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=news_csyy_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</a></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			  if rsnews("news_zz")<>"" then 
				dwt.out rsnews("news_zz")
			  else
				dwt.out rsnews("user_id")
			  end if
			 dwt.out "</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"" style=""color:0012FF"">"&rsnews("view_numb")
			 if isre then
                           sqlcsyy="select count(id) from csyy_re where news_id="&rsnews("id")
	                   set rscsyy=server.createobject("adodb.recordset")
	                       rscsyy.open sqlcsyy,conncs,1,1
                           if rscsyy.eof and rscsyy.bof then
                              dwt.out "/"
                           else
                              csyy_numb=conncs.Execute(sqlcsyy)(0)
                              dwt.out "/"&csyy_numb
                           end if
                           rscsyy.close
                           set rscsyy=nothing
                        end if
                        
			 dwt.out "</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsnews("news_date")&"</div></td>"
			 dwt.out  "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
'			 if rsnews("user_id")=0 or session("userid")=79 or session("userid")=1 or rsnews("user_id")=session("userid") then 
'				 dwt.out  "<a href='news.asp?classid="&classid&"&action=edit&ID="&rsnews("id")&"'>编辑</a>&nbsp;"
		 if  session("groupid")=3  then 
				 dwt.out "<a href='news_csyy.asp?action=del&ID="&rsnews("id")&"' onClick=""return confirm('确定要删除此内容吗？');"">删除</a>"
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
dim rsdel2,sqldel2
ID=request("ID")



	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from csyy_body where id="&id
	rsdel.open sqldel,conncsyy,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	'rsdel.close
	set rsdel=nothing  
	set rsdel2=server.createobject("adodb.recordset")
	sqldel="delete * from csyy_re where news_id="&id
	rsdel2.open sqldel,conncsyy,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	'rsdel.close
	set rsdel2=nothing  
end sub
dwt.out  "</body></html>"
sub search()
	dim sqlcj,rscj
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='news_csyy.asp'>" & vbCrLf
	Dwt.out "<a href=""news_csyy.asp?action=add&classid="&classid&""">添加</a>"
	dwt.out "&nbsp;&nbsp;搜索：" & vbCrLf
	Dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	Dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf
end sub








Call CloseConn
%>