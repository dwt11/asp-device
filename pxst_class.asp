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


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh
dim rsadd,sqladd,TrueIP,id,rsedit,sqledit,rsdel,sqldel
dim sqluser,rsuser,sqlcj,rscj
url="ghmanagement.asp"
dwt.out "<html>"
dwt.out "<head>"
dwt.out "<title>试题库分类管理</title>"
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function CheckAdd(){" & vbCrLf
 dwt.out " if(document.form1.class_name.value==''){" & vbCrLf
dwt.out "      alert('名称不能为空！');" & vbCrLf
dwt.out "   document.form1.class_name.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"


if Request("action")="add" then 
   call add
else
   if Request("action")="saveadd" then
      call saveadd
   else
	  if request("action")="edit" then 
	     call edit
	  else	 
	    if request("action")="saveedit" then
		    call saveedit
		else	
		    if request("action")="del" then
			   call del
			   'dwt.out"11111"
			else
			   call main 
			end if    
		end if 	
	  end if 	 
    end if  
end if 

sub add()
   '新增
   dwt.out"<form method='post' action='pxst_class.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>新增试题库分类</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 dwt.out"<strong>试题库分类名：</strong></td>"
	 dwt.out"<td width='80%' class='tdbg'><input name='class_name' type='text'></td>    </tr>   "
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='pxst_class.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from pxst_class" 
      rsadd.open sqladd,connpxjhzj,1,3
      rsadd.addnew
      rsadd("class_name")=ReplaceBadChar(Trim(Request("class_name")))
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
	
	  
end sub

sub main()
     	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>试题库分类管理</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
'用户管理首页
	dwt.out "<div class='x-toolbar'><div align=left><a href='pxst_class.asp?action=add'>添加分类</a></div></div>" & vbCrLf
 		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf

	  dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      dwt.out "<tr  class=""x-grid-header"">" 
      dwt.out "     <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
      dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center""><strong>试题库分类名</strong></div></td>"
     dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>操作</strong></div></td>"
      dwt.out "    </tr>"
      sqlbody="SELECT * from pxst_class "
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,connpxjhzj,1,1
      if rsbody.eof and rsbody.bof then 
           dwt.out "<p align=""center"">暂无内容</p>" 
      else
           record=rsbody.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsbody.PageSize = Cint(PgSz) 
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
           rsbody.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsbody.PageSize
           do while not rsbody.eof and rowcount>0
              
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center"">"&rsbody("class_name")&"</div></td>"
                  dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='pxst_class.asp?action=edit&ID="&rsbody("id")&"'>编辑</a>&nbsp;"
				 dwt.out "  <a href='pxst_class.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('确定要删除此试题库分类吗？删除后其相关的此分类内容将不会显示');"">删除</a></div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
		dwt.out "</table>"& vbCrLf
		call showpage1(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
       end if
 	dwt.out "</div>"  
      rsbody.close
       set rsbody=nothing
       
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from pxst_class where id="&id
   rsedit.open sqledit,connpxjhzj,1,1

   dwt.out"<form method='post' action='pxst_class.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>编辑试题分类</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>试题分类名称：</strong></td> "
	 dwt.out"<td width='80%' class='tdbg'>"
	 dwt.out"<input type='text' name='class_name' value="&rsedit("class_name")&"></td>    </tr> "
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from pxst_class where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connpxjhzj,1,3
rsedit("class_name")=ReplaceBadChar(Trim(Request("class_name")))
rsedit.update
rsedit.close
	
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from pxst_class where id="&id
rsdel.open sqldel,connpxjhzj,1,3
dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub

dwt.out "</body></html>"

Call CloseConn
%>