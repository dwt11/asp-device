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
<!--#include file="inc/md5.asp"-->


<%
dim sqljgtz,rsjgtz,title,record,pgsz,total,page,start,rowcount,url,ii,zxzz
dim rsadd,sqladd,jgtzid,rsedit,sqledit,scontent,rsdel,sqldel,tyzk,id
jgtzid=request("jgtzid")
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ����̨��-��������ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

if Request("action")="add" then call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del

sub add()
dim jgtzname,sql,rs
sql="SELECT * from jgtz where id="&jgtzid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
	jgtzname=rs("name")
rs.close
set rs=nothing
   response.write"<br><br><br><form method='post' action='jgtz_bj.asp' name='form1'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>��Ӽ���̨��  "&jgtzname&"   ��������</strong></div></td>    </tr>"
	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
     response.write"<input name='jgtz_bj_name' type='text' ></td></tr>"& vbCrLf
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>�����ͺţ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='jgtz_bj_xh' type='text'  size=""50""></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='jgtz_bj_sl'></td></tr> "
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ע��</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='jgtz_bj_bz'></td></tr> "
 
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'> <input name='jgtzid' type='hidden'  value='"&Trim(Request("jgtzid"))&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from jgtz_bj" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("jgtzid")=request("jgtzid")
	  rsadd("bj_name")=Trim(Request("jgtz_bj_name"))
      rsadd("bj_xh")=request("jgtz_bj_xh")
      rsadd("bj_sl")=request("jgtz_bj_sl")
      rsadd("bj_bz")=request("jgtz_bj_bz")

	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>history.go(-2);</Script>"
end sub

sub del()
  jgtzid=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from jgtz_bj where id="&jgtzid
  rsdel.open sqldel,connjg,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub saveedit()    
	  '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jgtz_bj where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
	  rsedit("bj_name")=Trim(Request("jgtz_bj_name"))
      rsedit("bj_xh")=request("jgtz_bj_xh")
      rsedit("bj_sl")=request("jgtz_bj_sl")
      rsedit("bj_bz")=request("jgtz_bj_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2);</Script>"
end sub



sub edit()
   dim jgtzname,sql,rs
sql="SELECT * from jgtz where id="&jgtzid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
	jgtzname=rs("name")
rs.close
set rs=nothing
id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from jgtz_bj where id="&id
   rsedit.open sqledit,connjg,1,1
   response.write"<br><br><br><form method='post' action='jgtz_bj.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�༭����̨��  "&jgtzname&"  ��������</strong></div></td>    </tr>"
	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
     response.write"<input name='jgtz_bj_name' type='text' value='"&rsedit("bj_name")&"'></td></tr>"& vbCrLf
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>�����ͺţ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='jgtz_bj_xh' type='text'  size=""50"" value='"&rsedit("bj_xh")&"'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='jgtz_bj_sl' value='"&rsedit("bj_sl")&"'></td></tr> "
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ע��</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='jgtz_bj_bz' value='"&rsedit("bj_bz")&"'></td></tr> "


	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

end sub


Call Closeconn
%>