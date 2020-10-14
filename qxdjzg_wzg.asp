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
dwt.out "<title>信息管理系统缺陷隐患登记整改管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
 
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


sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from qxdjzg where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connb,1,3
	on error resume next
	'rsedit("sscj")=Trim(Request("qxdj_sscj"))
	rsedit("wh")=request("qxdj_wh")
	rsedit("body")=Trim(request("qxdj_body"))
	rsedit("cxdate")=request("qxdj_cxdate")
	if request("qxdj_zgjg")=false then
	zgdate="lllll"
	else
	zgdate=request("qxdj_zgdate")
	end if   
	rsedit("zgdate")=zgdate
	rsedit("zgbody")=request("qxdj_zgbody")
	rsedit("zgjg")=request("qxdj_zgjg")
	rsedit("dbname")=request("qxdj_dbname")
	rsedit("update")=now()
	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from qxdjzg where id="&id
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
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑车间缺陷记录</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='qxdjzg.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px'  value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' >"& vbCrLf
	dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' type='hidden' name='qxdj_sscj' value='"&sscjh(rsedit("sscj"))&"'  >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >检修原因:</LABEL>"& vbCrLf
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
	dwt.out"				<LABEL style='WIDTH: 75px'>出现时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='qxdj_cxdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("cxdate")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>整改结果:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<select name='qxdj_zgjg' style='WIDTH: 175px' size='1'>"
	dwt.out"<option value='true'"
	if rsedit("zgjg") then dwt.out "selected"
	dwt.out ">已整改</option>"
    dwt.out"<option value='false'"
	if rsedit("zgjg")=false then dwt.out "selected"
	dwt.out ">未整改</option>"
    dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >缺陷内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <TEXTAREA class='x-form-textarea x-form-field' style='OVERFLOW: hidden; WIDTH: 175px; HEIGHT: 60px' name=qxdj_zgbody >"&rsedit("zgbody")&"</TEXTAREA>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>督办人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=qxdj_dbname value='"&rsedit("dbname")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>整改时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='qxdj_zgdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("zgdate")&"'>"
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
sub del()
	qxdjid=request("id")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from qxdjzg where id="&qxdjid
	rsdel.open sqldel,connb,1,3
	dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
	set rsdel=nothing  
end sub

sub main()
	
	sqlqxdj="SELECT * from qxdjzg where zgjg=0"
	if keys<>"" then 
		sqlqxdj=sqlqxdj&" where wh like '%"&keys&"%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then	
	   if sscjid=1000 then 
		sqlqxdj=sqlqxdj&" where isck"
		title="-厂控"
           else
          sqlqxdj=sqlqxdj&" and sscj="&sscjid
		title="-"&sscjh(sscjid)
           end if 
	end if 
	
	sqlqxdj=sqlqxdj&" ORDER BY sscj aSC,cxdate desc"
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>车间缺陷记录"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

	for sscji=1 to 5 
		sql="SELECT count(id) FROM qxdjzg WHERE sscj="&sscji&" and zgjg=0"
		numb=numb&sscjh_d(sscji)&"<span style='color:#006600;'>"&connb.Execute(sql)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
	next
	
	sql="SELECT count(id) FROM qxdjzg WHERE  zgjg=0"
	totall= "<span style='color:#006600;'>"&connb.Execute(sql)(0)&"</span>" 
	dwt.out "<div class='pre'>未整改:"&numb&"合计:"&totall&"</div>"& vbCrLf
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
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>出现年月</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>整改状态</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>整改结果</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>督办人</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>整改日期</div></td>"
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
			if rsqxdj("isck") then dwt.out " <span class='red'>厂控</span> "
			dwt.out sscjh_d(rsqxdj("sscj"))&"</div></td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"""
			if now()-rsqxdj("update")<7 then   dwt.out "bgcolor=""#FFFF00"""
			dwt.out ">"
			if rsqxdj("zgjg") then 
			   dwt.out searchH(uCase(rsqxdj("wh")),keys)
			else
			   dwt.out "<font color='#ff0000'>"&searchH(uCase(rsqxdj("wh")),keys)&"<font>"
			end if  
			   dwt.out "&nbsp;</td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("body")&"&nbsp;</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("cxdate")&"&nbsp;</div></td>"
			if rsqxdj("zgjg") then 
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">已整改</td>"
			else
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">未整改</td>"
			end if 
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsqxdj("zgbody")&"&nbsp;</td>"
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("dbname")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("zgdate")&"&nbsp;</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"
			'如果LEVELCLASS=7则显示设置厂控缺陷
			if session("levelclass")=7 and rsqxdj("isck")=false then dwt.out "<a href='qxdjzg.asp?id="&rsqxdj("id")&"&action=isck' onClick=""return confirm('确定设置为厂控缺陷吗？');"">厂控缺陷</a> "
			call editdel(rsqxdj("id"),rsqxdj("sscj"),"qxdjzg.asp?action=edit&id=","qxdjzg.asp?action=del&id=")
			
			dwt.out "</div></td></tr>"
			 RowCount=RowCount-1
          rsqxdj.movenext
		loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" or sscjid<>"" or request("allnochange")=1 then
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
	'if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "  <a href=""qxdjzg.asp?action=add"">添加记录</a>&nbsp;&nbsp;"
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
			dwt.out"<option value='qxdjzg_wzg.asp?sscj="&rscj("levelid")&"'"
	if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
			dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext	
		loop
		rscj.close
		set rscj=nothing
		dwt.out "<option value='qxdjzg.asp?sscj=1000'>厂控</option>"
		dwt.out "     </select>	" & vbCrLf
	dwt.out "<a href=qxdjzg.asp?allnochange=1>全部未整改</a>"
	dwt.out "</div></div></form>" & vbCrLf
end sub





Call CloseConn
%>