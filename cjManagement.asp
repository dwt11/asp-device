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
dim sqluser,rsuser
url="cjmanagement.asp"
response.write "<html>"
response.write "<head>"
response.write "<title>车间管理</title>"
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function CheckAdd(){" & vbCrLf
 response.write " if(document.form1.username.value==''){" & vbCrLf
response.write "      alert('用户名不能为空！');" & vbCrLf
response.write "   document.form1.username.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.password.value==''){" & vbCrLf
response.write "      alert('密码不能为空！');" & vbCrLf
response.write "  document.form1.password.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.password.value!=document.form1.password1.value){" & vbCrLf
response.write "      alert('两次输入的密码不一样！');" & vbCrLf
response.write "  document.form1.password.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.lxclassid.value==''){" & vbCrLf
response.write "      alert('未设置用户权限！');" & vbCrLf
 response.write "  document.form1.lxclassid.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
response.write " <tr class='topbg'>"
response.write "   <td height='22' colspan='2' align='center'><strong>车 间 管 理</strong></td>"
response.write "  </tr>  "

response.write " <tr class='tdbg'><td width='70' height='30'><strong>管理导航：</strong></td>"
response.write "    <td height='30'><a href='cjManagement.asp'>车间管理首页</a>&nbsp;|&nbsp;<a href='cjManagement.asp?action=add'>新增车间</a>    </td>"
response.write "  </tr>"
response.write "</table>"


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
			   'response.write"11111"
			else
			   call main 
			end if    
		end if 	
	  end if 	 
    end if  
end if 

sub add()
   '新增用户
   response.write"<form method='post' action='usermanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>新 增 用 户</strong></div></td>    </tr>"
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>用 户 名：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='username' type='text'></td>    </tr>   "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>密&nbsp;&nbsp;&nbsp;&nbsp;码：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>确认密码：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password1' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>权限设置： </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	response.write"<select name='lxclassid' size='1'>"
    response.write"<option selected>请选择权限分类</option>"
	response.write"<option value='1'>维修一车间</option>"
    response.write"<option value='2'>维修二车间</option>"
    response.write"<option value='3'>维修三车间</option>"
    response.write"<option value='4'>维修四车间</option>"
    response.write"<option value='5'>综合车间</option>"
    response.write"<option value='6'>计量车间</option>"
    response.write"<option value='7'>技术科</option>"
    response.write"<option value='8'>分厂领导</option>"
       response.write"<option value='9'>办公室</option>"
 response.write"</select></td></tr>  "   
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='usermanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '保存新增用户
   'set rsuser=server.createobject("adodb.recordset")
   'sqluser="select * from userid where username="&Request("username")
  ' rsuser.open sqluser,conn,1,1
  ' if rsuser.eof and rsuser.bof then 
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from userid" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
      rsadd("username")=ReplaceBadChar(Trim(Request("username")))
      rsadd("password")=md5(request("password"),16)
      rsadd("level")=ReplaceBadChar(Trim(request("lxclassid")))
      rsadd("dldate")=now()
      TrueIP=Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
      If TrueIP = "" Then TrueIP = Request.ServerVariables("REMOTE_ADDR")
	  rsadd("dlip")=TrueIP
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>window.alert('用户添加成功');location.href='usermanagement.asp';</Script>"
	'else
'  
  
   'end if 
   'rsuser.close
  ' set rsuser=nothing
	
	  
end sub

sub main()
     '用户管理首页
	  response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      response.write "<tr class=""title"">" 
      response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
      response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center""><strong>用户名</strong></div></td>"
      response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>用户等级</strong></div></td>"
      response.write "      <td width=""14%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>最后登录时间</strong></div></td>"
      response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>最后登录IP</strong></div></td>"
      response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>登录次数</strong></div></td>"
      response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>操作</strong></div></td>"
      response.write "    </tr>"
      sqlbody="SELECT * from userid "
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,conn,1,1
      if rsbody.eof and rsbody.bof then 
           response.write "<p align=""center"">暂无内容</p>" 
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
                 xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
                 response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center"">"&rsbody("username")&"</div></td>"
                  sqllevel="SELECT * from levelname where levelid="&rsbody("levelid")
                 set rslevel=server.createobject("adodb.recordset")
                 rslevel.open sqllevel,conn,1,1
                 if rslevel.eof and rslevel.bof then 
                     response.write "   <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">暂无内容</div></td>" 
                 else 
                     response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rslevel("levelname")&"</div></td>"
                 end if
                 rslevel.close
                 set rslevel=nothing
                 response.write "      <td width=""14%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dldate")&"</div></td>"
                 response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dlip")&"</div></td>"
                 response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dlcs")&"</div></td>"
                  response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='usermanagement.asp?action=edit&ID="&rsbody("id")&"'>编辑</a>&nbsp;"
				 if rsbody("levelid")>0 then response.write "  <a href='usermanagement.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('确定要删除此用户吗？');"">删除</a></div></td>"
                 response.write "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
       end if
       rsbody.close
       set rsbody=nothing
        conn.close
        set conn=nothing
        response.write "</table>"
       call showpage1(page,url,total,record,PgSz)
end sub

sub edit()
     '编辑用户
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from userid where id="&id
   rsedit.open sqledit,conn,1,1

   response.write"<form method='post' action='usermanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>编 辑 用 户</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>用 户 名：</strong></td>"
   if session("level")=0 then 
	if rsedit("level")=0 then 
		  response.write"<td width='88%' class='tdbg'><input name='username' type='text' disabled='true'  value='"&rsedit("username")&"'></td>    </tr>   "
 else
	  response.write"<td width='88%' class='tdbg'><input name='username' type='text' value='"&rsedit("username")&"'></td>    </tr>   "

	 end if 
 else 
		  response.write"<td width='88%' class='tdbg'><input name='username' type='text' disabled='true' value='"&rsedit("username")&"'></td>    </tr>   "
end if 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>密&nbsp;&nbsp;&nbsp;&nbsp;码：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password1' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>确认密码：</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password' ></td>    </tr> "
	 if session("level")=0 then 
	   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>权限设置： </strong></td>"      
       response.write"<td width='88%' class='tdbg'>"
	   
	   	if rsedit("level")=0 then 
	    response.write"<select name='lxclassid' size='1' disabled='disabled'  onChange=""redirect(this.options.selectedIndex)"">"
     else
			     response.write"<select name='lxclassid' size='1' onChange=""redirect(this.options.selectedIndex)"">"
end if 
       response.write"<option"
	  if rsedit("level")="" then response.write "selected" 
	      response.write">请选择权限分类</option>"
	      response.write"<option value='1' "
	   if rsedit("level")=1 then response.write "selected"
	 response.write">维修一车间</option>"
    response.write"<option value='2'"
	if rsedit("level")=2 then response.write "selected"
    response.write" >维修二车间</option>"
    response.write"<option value='3'"
	if rsedit("level")=3 then response.write "selected"
    response.write">维修三车间</option>"
    response.write"<option value='4'"
    if rsedit("level")=4 then response.write "selected"
	response.write">维修四车间</option>"
    response.write"<option value='5'"
	if rsedit("level")=5 then response.write "selected"
	response.write">综合车间</option>"
    response.write"<option value='6'"
	if rsedit("level")=6 then response.write "selected"
	response.write">计量车间</option>"
    response.write"<option value='7'"
	if rsedit("level")=7 then response.write "selected"
	response.write">技术科</option>"
    response.write"<option value='8'"
	if rsedit("level")=8 then response.write "selected"
	response.write">分厂领导</option>"
       response.write"<option value='0'"
	if rsedit("level")=0 then response.write "selected"
	response.write">超级管理员</option>"
if rsedit("level")=9 then response.write "selected"
	response.write">办公室</option>"
    response.write"</select></td></tr>  "
	else 
	 response.write" <input type='hidden' name='lxclassid' value='"&rsedit("level")&"'>"
	end if    
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from userid where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conn,1,3
rsedit("username")=ReplaceBadChar(Trim(Request("username")))
rsedit("password")=md5(request("password"),16)
rsedit("level")=ReplaceBadChar(Trim(request("lxclassid")))
rsedit("dldate")=now()
TrueIP=Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
If TrueIP = "" Then TrueIP = Request.ServerVariables("REMOTE_ADDR")
rsedit("dlip")=TrueIP
rsedit.update
rsedit.close
	if session("level")=0 then 
        response.write"<Script Language=Javascript>window.alert('用户编辑成功');location.href='usermanagement.asp';</Script>"
    else
	  response.write"<Script Language=Javascript>window.alert('用户编辑成功');history.back()</Script>"
	 end if 
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from userid where id="&id
rsdel.open sqldel,conn,1,3
response.write"<Script Language=Javascript>window.alert('删除用户成功');location.href='usermanagement.asp';</Script>"
'rsdel.close
set rsdel=nothing  

end sub

response.write "</body></html>"

Call CloseConn
%>