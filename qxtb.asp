<%@language=vbscript codepage=936 %>
<%
Option Explicit
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
dim sqlqxtb,rsqxtb,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
url="qxtb.asp"
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统缺陷整改通知管理页</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
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
	dwt.out "<br><br><form method='post' action='qxtb.asp' name='form1'>"
	dwt.out "<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out "<tr class='title'><td height='22' colspan='2'>"
	dwt.out "<div align='center'><strong>添加缺陷整改通知</strong></div></td>    </tr>"
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>标&nbsp;&nbsp;题：</strong></td>"
	dwt.out "<td width='88%' class='tdbg'><input name='qxtb_title' type='text'></td>    </tr>   "
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>发布人：</strong></td> "
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input type='text' name='qxtb_zz' ></td>    </tr> "
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>日&nbsp;&nbsp;期：</strong></td> "
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='qxtb_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out "</td></tr>"& vbCrLf
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=qxtb_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
	dwt.out "</iframe>  <input type='hidden' name='qxtb_body' value=''>"
	dwt.out "</td></tr>  "   
	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='qxtb.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	

sub saveadd()    
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from scgl_qxtb" 
	rsadd.open sqladd,connscgl,1,3
	rsadd.addnew
	rsadd("qxtb_title")=ReplaceBadChar(Trim(Request("qxtb_title")))
	rsadd("qxtb_zz")=request("qxtb_zz")
	rsadd("qxtb_body")=Trim(request("qxtb_body"))
	rsadd("qxtb_date")=request("qxtb_date")
	rsadd("userid")=session("userid")
	rsadd.update
	rsadd.close
	set rsadd=nothing
	dwt.out "<Script Language=Javascript>location.href='qxtb.asp';</Script>"
end sub

sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from scgl_qxtb where id="&id
	rsedit.open sqledit,connscgl,1,1
	dwt.out "<br><br><form method='post' action='qxtb.asp' name='form1'>"
	dwt.out "<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out "<tr class='title'><td height='22' colspan='2'>"
	dwt.out "<div align='center'><strong>编辑缺陷整改通知</strong></div></td>    </tr>"
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>标&nbsp;&nbsp;题：</strong></td>"
	dwt.out "<td width='88%' class='tdbg'><input name='qxtb_title' type='text' value='"&rsedit("qxtb_title")&"'></td>    </tr>   "
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>发布人：</strong></td> "
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='qxtb_zz' type='text' value='"&rsedit("qxtb_zz")&"'></td>    </tr> "
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>日&nbsp;&nbsp;期：</strong></td> "
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='qxtb_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("qxtb_date")&"'>"
	dwt.out "</td></tr>"& vbCrLf
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
	dwt.out "<td width='88%' class='tdbg'>"
	scontent=rsedit("qxtb_body")
	dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=qxtb_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
	dwt.out "</iframe><textarea name='qxtb_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
	dwt.out "</td></tr>  "   
	dwt.out "<tr> <td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='qxtb.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
	rsedit.close
	set rsedit=nothing
end sub

sub saveedit()
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from scgl_qxtb where ID="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connscgl,1,3
	rsedit("qxtb_title")=ReplaceBadChar(Trim(Request("qxtb_title")))
	rsedit("qxtb_zz")=request("qxtb_zz") 
	rsedit("qxtb_body")=Trim(request("qxtb_body"))
	rsedit("qxtb_date")=request("qxtb_date")
	rsedit.update
	rsedit.close
	dwt.out "<Script Language=Javascript>history.go(-2);</Script>"
end sub


sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>缺陷整改通知反馈</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	dwt.out "	  <div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "	  <a href=qxtb.asp?action=add>   添加整改通知</a>"
	dwt.out "	  </div></div>"
	sqlqxtb="SELECT * from scgl_qxtb ORDER BY id DESC"
	set rsqxtb=server.createobject("adodb.recordset")
	rsqxtb.open sqlqxtb,connscgl,1,1
	if rsqxtb.eof and rsqxtb.bof then 
		dwt.message  "<p align='center'>未添加内容</p>" 
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out  "<td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out  "<td class='x-td'><DIV class='x-grid-hd-text'>缺陷整改通知标题</div></td>"
		dwt.out  "<td class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out  "<td class='x-td'><DIV class='x-grid-hd-text'>发布时间</div></td>"
		dwt.out  "<td class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
		dwt.out  "</tr>"
		record=rsqxtb.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
		end if 
		rsqxtb.PageSize = Cint(PgSz) 
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
		rsqxtb.absolutePage = page
		start=PgSz*Page-PgSz+1
		DIM XH,XH_ID
		rowCount = rsqxtb.PageSize
		do while not rsqxtb.eof and rowcount>0
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rsqxtb("id")&"</div></td>"
			dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""30%""><a href=qxtb_view.asp?id="&rsqxtb("id")&" target=_blank>"&rsqxtb("qxtb_title")&"</a></td>"
			dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqxtb("qxtb_zz")&"</div></td>"
			dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&year(rsqxtb("qxtb_date"))&"年"&month(rsqxtb("qxtb_date"))&"月"&day(rsqxtb("qxtb_date"))&"日</div></td>"
			dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			if session("levelclass")<>0 and session("levelclass")<>9 then 
				dim rsqxtb_fk,sqlqxtb_fk
				set rsqxtb_fk=server.createobject("adodb.recordset")
				sqlqxtb_fk="select * from scgl_qxtb_fk where qxtb_fk_sscj="&session("levelclass")&" and qxtb_id="&rsqxtb("id")
				rsqxtb_fk.open sqlqxtb_fk,connb,1,1
				if rsqxtb_fk.eof and rsqxtb_fk.bof then 
					dwt.out  "<a href='qxtb_fk.asp?action=add&qxtb_fk_sscj="&session("levelclass")&"&qxtb_id="&rsqxtb("id")&"'>添加反馈</a>&nbsp;"
				else
					dwt.out  "<a href='qxtb_fk.asp?action=edit&qxtb_fk_sscj="&session("levelclass")&"&qxtb_id="&rsqxtb("id")&"'>编辑反馈</a>&nbsp;"
					if session("level")=0 then dwt.out  "<a href='qxtb_fk.asp?action=del&qxtb_fk_sscj="&session("levelclass")&"&qxtb_id="&rsqxtb("id")&"' onClick=""return confirm('确定要删除此反馈吗？');"">删除反馈</a>"
				end if 
				rsqxtb_fk.close
				set rsqxtb_fk=nothing
			end if 
			'dwt.out rsqxtb("userid") &"dfgdfg"&session("userid")
if session("levelclass")=0 or session("levelclass")=9 or session("userid")=rsqxtb("userid") then dwt.out  "<a href='qxtb.asp?action=edit&ID="&rsqxtb("id")&"'>编辑</a>&nbsp;  <a href='qxtb.asp?action=del&ID="&rsqxtb("id")&"' onClick=""return confirm('确定要删除此缺陷整改通知吗？');"">删除</a>"
			dwt.out  "</div></td>"
			dwt.out  "    </tr>"
			RowCount=RowCount-1
		rsqxtb.movenext
		loop
		dwt.out  "</table>"
		call showpage1(page,url,total,record,PgSz)
	    dwt.out "</div>"
	end if
	dwt.out "</div>"
	rsqxtb.close
	set rsqxtb=nothing
	conn.close
	set conn=nothing
end sub



sub del()
	ID=request("ID")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from scgl_qxtb where id="&id
	rsdel.open sqldel,connscgl,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	set rsdel=nothing  
end sub
dwt.out  "</body></html>"
Call CloseConn
%>