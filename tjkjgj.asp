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
dim sqltjkjgj,rstjkjgj,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
url="tjkjgj.asp"

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统科技信息管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf

response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

if Request("action")="add" then call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del
if request("action")="" then call main 

sub add()
   '新增用户
   response.write"<br><form method='post' action='tjkjgj.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>添 加 科 技 信 息  </strong></div></td>    </tr>"
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>标&nbsp;&nbsp;题：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='tjkjgj_title' type='text'></td>    </tr>   "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作者：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='text' name='tjkjgj_zz' ></td>    </tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=tjkjgj_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe>  <input type='hidden' name='tjkjgj_body' value=''>"
    response.write"</td></tr>  "   
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='tjkjgj.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from tjkjgj" 
      rsadd.open sqladd,conne,1,3
      rsadd.addnew
      rsadd("tjkjgj_title")=ReplaceBadChar(Trim(Request("tjkjgj_title")))
      rsadd("tjkjgj_zz")=request("tjkjgj_zz")
      rsadd("tjkjgj_body")=Trim(request("tjkjgj_body"))
      rsadd("tjkjgj_date")=year(now())&"-"&month(now())&"-"&day(now())
      
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='tjkjgj.asp';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from tjkjgj where id="&id
   rsedit.open sqledit,conne,1,1

   response.write"<br><br><br><form method='post' action='tjkjgj.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>编 辑 科 技 信 息</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
   response.write"<strong>标&nbsp;&nbsp;题：</strong></td>"
   response.write"<td width='88%' class='tdbg'><input name='tjkjgj_title' type='text' value='"&rsedit("tjkjgj_title")&"'></td>    </tr>   "
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>作者：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input name='tjkjgj_zz' type='text' value='"&rsedit("tjkjgj_zz")&"'></td>    </tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	scontent=rsedit("tjkjgj_body")
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=tjkjgj_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe><textarea name='tjkjgj_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    response.write"</td></tr>  "   

	 
    response.write"<tr> <td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='tjkjgj.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from tjkjgj where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conne,1,3
rsedit("tjkjgj_title")=ReplaceBadChar(Trim(Request("tjkjgj_title")))
rsedit("tjkjgj_zz")=request("tjkjgj_zz") 
rsedit("tjkjgj_body")=Trim(request("tjkjgj_body"))
rsedit("tjkjgj_date")=year(now())&"-"&month(now())&"-"&day(now())
rsedit.update
rsedit.close
	  response.write"<Script Language=Javascript>history.go(-2);</Script>"
	
end sub


sub main()
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>科技信息管理页</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='30'><a href='tjkjgj.asp'>科技信息管理首页</a>"
'if session("level")=0 or session("level")=9 then 
response.write "&nbsp;|&nbsp;<a href='tjkjgj.asp?action=add'>添加科技信息</a>"
response.write "    </td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

sqltjkjgj="SELECT * from tjkjgj ORDER BY id DESC"
set rstjkjgj=server.createobject("adodb.recordset")
rstjkjgj.open sqltjkjgj,conne,1,1
if rstjkjgj.eof and rstjkjgj.bof then 
message("未添加科技信息")
else
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><div align=""center""><strong>标    题</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>作者</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>发布时间</strong></div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
           record=rstjkjgj.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rstjkjgj.PageSize = Cint(PgSz) 
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
           rstjkjgj.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rstjkjgj.PageSize
           do while not rstjkjgj.eof and rowcount>0
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rstjkjgj("id")&"</div></td>"
                 response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=tjkjgj_view.asp?id="&rstjkjgj("id")&" target=_blank>"&rstjkjgj("tjkjgj_title")&"</a></td>"
                 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rstjkjgj("tjkjgj_zz")&"&nbsp;</div></td>"
                 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rstjkjgj("tjkjgj_date")&"</div></td>"
                 response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='tjkjgj.asp?action=edit&ID="&rstjkjgj("id")&"'>编辑</a>"
				 if session("level")=0 or session("level")=9 then response.write "&nbsp;<a href='tjkjgj.asp?action=del&ID="&rstjkjgj("id")&"' onClick=""return confirm('确定要删除此信息吗？');"">删除</a>"
				 response.write "&nbsp; </div></td>"
                 response.write "    </tr>"
                 RowCount=RowCount-1
          rstjkjgj.movenext
          loop
        response.write "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rstjkjgj.close
       set rstjkjgj=nothing
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from tjkjgj where id="&id
rsdel.open sqldel,conne,1,3
response.write"<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub







response.write "</body></html>"

Call CloseConn
%>