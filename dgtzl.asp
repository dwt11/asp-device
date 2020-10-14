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
classid=request("classid")
indexid=request("indexid")
if classid<>"" then classname=conndgt.Execute("SELECT class_name FROM dgtzl_class WHERE id="&classid)(0)
if classid="" then  classid=1

if indexid<>"" then classname=conndgt.Execute("SELECT class_name FROM dgtzl_index WHERE id="&indexid)(0)

keys=trim(request("keyword")) 
    url=geturl
   
dwt.pagetop classname
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
	dwt.out "<form method='post' action='dgtzl.asp' name='form1'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>是否记录查看人员:</DIV></div></LABEL>"& vbCrLf
	dwt.out"				<DIV >"& vbCrLf
	dwt.out"				  <div align=left><INPUT name=isviewd type=checkbox></DIV>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>标题:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=news_title>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>分类:</DIV></LABEL>"& vbCrLf
		dwt.out"				<DIV class=x-form-element align=left>"& vbCrLf
		'dwt.out "<select name='news_class' onChange=""redirect(this.options.selectedIndex)"">"
		'dwt.out "<option value='0'>请选择分类</option>"& vbCrLf
		
		
		
				
	
	if classid=1 then 
	  if indexid="" then  indexid=1
	dwt.out "<input name='ssbz' type='hidden' value='"&indexid&"'  />"
	dwt.out "<input name='news_class' type='hidden' value='1'  />"
	classname=conndgt.Execute("SELECT class_name FROM dgtzl_index WHERE id="&indexid)(0)
end if 
	if classid<>1 then 
	dwt.out "<input name='ssbz' type='hidden' value='0'  />"
	dwt.out "<input name='news_class' type='hidden' value='"&classid&"'  />"
	end if 
	
		dwt.out classname& vbCrLf
		'dwt.out"</select>"
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
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>时间:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
dwt.out"<input name='news_date' style='WIDTH: 175px'  disabled='disabled'  value='"&NOW()&"'>"	
dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>内容:</DIV></LABEL>"& vbCrLf
	dwt.out"				<DIV style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=news_body&style=s_blue&originalfilename=d_originalfilename &savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='550' HEIGHT='350'>"
	 dwt.out "</iframe>  <input type='hidden' name='news_body' value=''>"	

	
	 dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	
	
	
	
	
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'><div align=center>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>      <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </div></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
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
      sqladd="select * from dgtzl_body" 
      rsadd.open sqladd,conndgt,1,3
      rsadd.addnew
      rsadd("news_title")=ReplaceBadChar(Trim(Request("news_title")))
	  rsadd("index")=request("ssbz")
      rsadd("user_id")=request("user_id")
      rsadd("news_body")=Trim(request("news_body"))
      rsadd("news_date")=NOW()
      rsadd("news_class")=request("news_class")
      if request("isviewd")="on" then rsadd("isviewd")=true
      rsadd.update
      rsadd.close
	dwt.savesl conndgt.Execute("SELECT class_name FROM dgtzl_class WHERE id="&request("news_class"))(0) ,"添加",ReplaceBadChar(Trim(Request("news_title")))
      set rsadd=nothing
	  dwt.out "<Script Language=Javascript>location.href='dgtzl.asp?classid="&request("news_class")&"';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from dgtzl_body where id="&id
   rsedit.open sqledit,conndgt,1,1
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
	dwt.out "<form method='post' action='dgtzl.asp' name='form1' >"
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
		dwt.out"				<LABEL style='WIDTH: 105px' ><div align=right>分类:</DIV></LABEL>"& vbCrLf
		dwt.out"				<DIV class=x-form-element align=left>"& vbCrLf
		
		if classid=1 then 
	  if indexid="" then  indexid=1
	dwt.out "<input name='ssbz' type='hidden' value='"&rsedit("index")&"'  />"
	dwt.out "<input name='news_class' type='hidden' value='1'  />"
	classname=conndgt.Execute("SELECT class_name FROM dgtzl_index WHERE id="&rsedit("index"))(0)
end if 
	if classid<>1 then 
	dwt.out "<input name='ssbz' type='hidden' value='0'  />"
	dwt.out "<input name='news_class' type='hidden' value='"&classid&"'  />"
	end if 
	
		dwt.out classname& vbCrLf
		'dwt.out"</select>"
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
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 105px'><div align=right>时间:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
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
	sqledit="select * from dgtzl_body where ID="&ReplaceBadChar(Trim(request("ID")))
	
	rsedit.open sqledit,conndgt,1,3
	rsedit("news_title")=ReplaceBadChar(Trim(Request("news_title")))
	rsedit("news_body")=Trim(request("news_body"))
	rsedit("index")=request("ssbz")
	rsedit("news_date")=NOW()
	rsedit("news_class")=request("news_class")
      if request("isviewd")="on" then rsedit("isviewd")=true
	rsedit.update
	rsedit.close
	dwt.savesl conndgt.Execute("SELECT class_name FROM dgtzl_class WHERE id="&request("news_class"))(0) ,"编辑",ReplaceBadChar(Trim(Request("news_title")))
	dwt.out "<Script Language=Javascript>history.go(-2);</Script>"
	
end sub


sub main()
dim sqlcsyy,rscsyy,csyy_numb
	if keys<>"" then 
		sqlbody="SELECT * from dgtzl_body  where news_title like '%" &keys& "%'  and news_class="&classid
		title="-搜索 "&keys
	else
	    	sqlbody="SELECT * from dgtzl_body WHERE news_class="&classid
	
	end if 
	if indexid<>0 then
	sqlbody=sqlbody&"and index="&indexid
	end if
	sqlbody=sqlbody&" order by ID desc"

    
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&classname&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
    
	
		search()
	set rsnews=server.createobject("adodb.recordset")
	rsnews.open sqlbody,conndgt,1,1
	if rsnews.eof and rsnews.bof then 
		dwt.out  message ("<p align='center'>未添加内容</p>" )
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out  "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>分类</div></td>"
		dwt.out  "      <td class='x-td'><DIV class='x-grid-hd-text'>标题</div></td>"
		dwt.out  "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out  "      <td  class='x-td'><DIV class='x-grid-hd-text'>浏览"
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
          dim classname11
		  if rsnews("news_class")=1 then 
		  'classname11=conndgt.Execute("SELECT class_name FROM dgtzl_index WHERE id="&rsnews("index"))(0)
		  
sqld="SELECT class_name FROM dgtzl_index WHERE id="&rsnews("index")
		set rsd=server.createobject("adodb.recordset")
		rsd.open sqld,conndgt,1,1
		if rsd.eof and rsd.eof then 
			classname11="无"
		else
			classname11=rsd("class_name")
		end if 	





end if 
		   'classname11="党委"
		  if rsnews("news_class")=2 then  classname11="工会"
		  if rsnews("news_class")=3 then  classname11="团委"

			 dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center"">"&classname11&"</div></td>"
			
			if classid=1 then 
			 dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=/dw/view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</a></td>"
			 else
			 dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=/dgtzl_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</a></td>"
			 end if 
			
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			  if rsnews("user_id")=0 then 
				dwt.out rsnews("news_zz")
			  else
				dwt.out usernameh(rsnews("user_id"))
			  end if
			 
			 dwt.out "</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"" style=""color:0012FF"">"&rsnews("view_numb")
                        
			 dwt.out "</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsnews("news_date")&"</div></td>"
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			 if rsnews("user_id")=0 or session("userid")=79 or session("userid")=1 or rsnews("user_id")=session("userid") then 
				 dwt.out  "<a href='dgtzl.asp?classid="&classid&"&action=edit&ID="&rsnews("id")&"'>编辑</a>&nbsp;"
				 dwt.out "<a href='dgtzl.asp?action=del&ID="&rsnews("id")&"' onClick=""return confirm('确定要删除此内容吗？');"">删除</a>"
			 end if 
			 dwt.out  "&nbsp; </div></td>"
			 dwt.out  "    </tr>"
			 RowCount=RowCount-1
	  rsnews.movenext
	  loop
	dwt.out  "</table>"
   IF indexid="" THEN
     call showpage(page,url,total,record,PgSz)
    ELSE
call showpage(page,url,total,record,PgSz)
END IF 
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


	sqledit="select * from dgtzl_body where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,conndgt,1,1

	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from dgtzl_body where id="&id
	rsdel.open sqldel,conndgt,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	set rsdel=nothing  

end sub

sub search()
	dim sqlcj,rscj
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='dgtzl.asp'>" & vbCrLf
	
	dim urlaction
	if classid<>1 then urlaction="classid="&classid
	if classid=1 then urlaction="indexid="&indexid
	if displaypagelevelh(session("groupid"),1,session("pagelevelid"))  then Dwt.out "<a href=""dgtzl.asp?action=add&"&urlaction&""">添加</a>"
	dwt.out "&nbsp;&nbsp;搜索：" & vbCrLf
	Dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	'Dwt.out "按组用户查看："
	Dwt.out "<select name='classid'>" & vbCrLf
	'Dwt.out "<option value=''>按组跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from dgtzl_class"& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conndgt,1,1
		do while not rscj.eof
			Dwt.out"<option value="&rscj("id")
			if cint(classid)=rscj("id") then Dwt.out" selected"
			Dwt.out">"&rscj("class_name")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		Dwt.out "</select>" & vbCrLf
	Dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf
end sub







dwt.out  "</body></html>"

Call CloseConn
%>