<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
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
response.write "<title>�������</title>"
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function CheckAdd(){" & vbCrLf
 response.write " if(document.form1.username.value==''){" & vbCrLf
response.write "      alert('�û�������Ϊ�գ�');" & vbCrLf
response.write "   document.form1.username.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.password.value==''){" & vbCrLf
response.write "      alert('���벻��Ϊ�գ�');" & vbCrLf
response.write "  document.form1.password.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.password.value!=document.form1.password1.value){" & vbCrLf
response.write "      alert('������������벻һ����');" & vbCrLf
response.write "  document.form1.password.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.lxclassid.value==''){" & vbCrLf
response.write "      alert('δ�����û�Ȩ�ޣ�');" & vbCrLf
 response.write "  document.form1.lxclassid.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
response.write " <tr class='topbg'>"
response.write "   <td height='22' colspan='2' align='center'><strong>�� �� �� ��</strong></td>"
response.write "  </tr>  "

response.write " <tr class='tdbg'><td width='70' height='30'><strong>��������</strong></td>"
response.write "    <td height='30'><a href='cjManagement.asp'>���������ҳ</a>&nbsp;|&nbsp;<a href='cjManagement.asp?action=add'>��������</a>    </td>"
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
   '�����û�
   response.write"<form method='post' action='usermanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�� �� �� ��</strong></div></td>    </tr>"
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>�� �� ����</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='username' type='text'></td>    </tr>   "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�룺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ȷ�����룺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password1' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>Ȩ�����ã� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	response.write"<select name='lxclassid' size='1'>"
    response.write"<option selected>��ѡ��Ȩ�޷���</option>"
	response.write"<option value='1'>ά��һ����</option>"
    response.write"<option value='2'>ά�޶�����</option>"
    response.write"<option value='3'>ά��������</option>"
    response.write"<option value='4'>ά���ĳ���</option>"
    response.write"<option value='5'>�ۺϳ���</option>"
    response.write"<option value='6'>��������</option>"
    response.write"<option value='7'>������</option>"
    response.write"<option value='8'>�ֳ��쵼</option>"
       response.write"<option value='9'>�칫��</option>"
 response.write"</select></td></tr>  "   
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='usermanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '���������û�
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
	  response.write"<Script Language=Javascript>window.alert('�û���ӳɹ�');location.href='usermanagement.asp';</Script>"
	'else
'  
  
   'end if 
   'rsuser.close
  ' set rsuser=nothing
	
	  
end sub

sub main()
     '�û�������ҳ
	  response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      response.write "<tr class=""title"">" 
      response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
      response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center""><strong>�û���</strong></div></td>"
      response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�û��ȼ�</strong></div></td>"
      response.write "      <td width=""14%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����¼ʱ��</strong></div></td>"
      response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����¼IP</strong></div></td>"
      response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��¼����</strong></div></td>"
      response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
      response.write "    </tr>"
      sqlbody="SELECT * from userid "
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,conn,1,1
      if rsbody.eof and rsbody.bof then 
           response.write "<p align=""center"">��������</p>" 
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
                     response.write "   <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">��������</div></td>" 
                 else 
                     response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rslevel("levelname")&"</div></td>"
                 end if
                 rslevel.close
                 set rslevel=nothing
                 response.write "      <td width=""14%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dldate")&"</div></td>"
                 response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dlip")&"</div></td>"
                 response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dlcs")&"</div></td>"
                  response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='usermanagement.asp?action=edit&ID="&rsbody("id")&"'>�༭</a>&nbsp;"
				 if rsbody("levelid")>0 then response.write "  <a href='usermanagement.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('ȷ��Ҫɾ�����û���');"">ɾ��</a></div></td>"
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
     '�༭�û�
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from userid where id="&id
   rsedit.open sqledit,conn,1,1

   response.write"<form method='post' action='usermanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�� �� �� ��</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>�� �� ����</strong></td>"
   if session("level")=0 then 
	if rsedit("level")=0 then 
		  response.write"<td width='88%' class='tdbg'><input name='username' type='text' disabled='true'  value='"&rsedit("username")&"'></td>    </tr>   "
 else
	  response.write"<td width='88%' class='tdbg'><input name='username' type='text' value='"&rsedit("username")&"'></td>    </tr>   "

	 end if 
 else 
		  response.write"<td width='88%' class='tdbg'><input name='username' type='text' disabled='true' value='"&rsedit("username")&"'></td>    </tr>   "
end if 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�룺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password1' ></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ȷ�����룺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='password' name='password' ></td>    </tr> "
	 if session("level")=0 then 
	   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>Ȩ�����ã� </strong></td>"      
       response.write"<td width='88%' class='tdbg'>"
	   
	   	if rsedit("level")=0 then 
	    response.write"<select name='lxclassid' size='1' disabled='disabled'  onChange=""redirect(this.options.selectedIndex)"">"
     else
			     response.write"<select name='lxclassid' size='1' onChange=""redirect(this.options.selectedIndex)"">"
end if 
       response.write"<option"
	  if rsedit("level")="" then response.write "selected" 
	      response.write">��ѡ��Ȩ�޷���</option>"
	      response.write"<option value='1' "
	   if rsedit("level")=1 then response.write "selected"
	 response.write">ά��һ����</option>"
    response.write"<option value='2'"
	if rsedit("level")=2 then response.write "selected"
    response.write" >ά�޶�����</option>"
    response.write"<option value='3'"
	if rsedit("level")=3 then response.write "selected"
    response.write">ά��������</option>"
    response.write"<option value='4'"
    if rsedit("level")=4 then response.write "selected"
	response.write">ά���ĳ���</option>"
    response.write"<option value='5'"
	if rsedit("level")=5 then response.write "selected"
	response.write">�ۺϳ���</option>"
    response.write"<option value='6'"
	if rsedit("level")=6 then response.write "selected"
	response.write">��������</option>"
    response.write"<option value='7'"
	if rsedit("level")=7 then response.write "selected"
	response.write">������</option>"
    response.write"<option value='8'"
	if rsedit("level")=8 then response.write "selected"
	response.write">�ֳ��쵼</option>"
       response.write"<option value='0'"
	if rsedit("level")=0 then response.write "selected"
	response.write">��������Ա</option>"
if rsedit("level")=9 then response.write "selected"
	response.write">�칫��</option>"
    response.write"</select></td></tr>  "
	else 
	 response.write" <input type='hidden' name='lxclassid' value='"&rsedit("level")&"'>"
	end if    
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
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
        response.write"<Script Language=Javascript>window.alert('�û��༭�ɹ�');location.href='usermanagement.asp';</Script>"
    else
	  response.write"<Script Language=Javascript>window.alert('�û��༭�ɹ�');history.back()</Script>"
	 end if 
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from userid where id="&id
rsdel.open sqldel,conn,1,3
response.write"<Script Language=Javascript>window.alert('ɾ���û��ɹ�');location.href='usermanagement.asp';</Script>"
'rsdel.close
set rsdel=nothing  

end sub

response.write "</body></html>"

Call CloseConn
%>