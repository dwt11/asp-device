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

<%
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql

'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>���̨�˹���ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checkadd(){" & vbCrLf
response.write "  if(document.form1.zjtz_name.value==''){" & vbCrLf
response.write "      alert('��������δ��д��');" & vbCrLf
response.write "  document.form1.zjtz_name.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if Request("action")="" then call mainclass '����
if Request("action")="addclass" then call addclass     '���ӷ���
if Request("action")="saveaddclass" then call saveaddclass     '�������
if Request("action")="editclass" then call editclass   '�༭����
if Request("action")="saveeditclass" then call saveeditclass  '�༭�������
if Request("action")="delclass" then call delclass     'ɾ��������Ϣ



sub addclass()'��ӷ���
   response.write"<br><br><br><form method='post' action='zjtz_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>���̨�ˣ����������</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    response.write"<td width='80%' class='tdbg'>"
       response.write"<input name='zjtz_name' type='text'></td></tr>"& vbCrLf
	   
	   
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveaddclass'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveaddclass()    
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from class" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
       on error resume next
      rsadd("name")=request("zjtz_name")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='zjtz_class.asp';</Script>"
end sub



sub saveeditclass()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connzj,1,3
      rsedit("name")=Trim(Request("zjtz_name"))
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub delclass()
dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from class where id="&request("id")
  rsdel.open sqldel,connzj,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub



sub editclass()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from class where id="&id
   rsedit.open sqledit,connzj,1,1
   response.write"<br><br><br><form method='post' action='zjtz_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�༭�ӷ�������</strong></div></td>    </tr>"
     
     response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ӷ������ƣ� </strong></td>"   & vbCrLf   
     response.write"<td width='80%' class='tdbg'><input name='zjtz_name' type='text' value='"&rsedit("name")&"'></td></tr>"& vbCrLf

		response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveeditclass'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
end sub


Sub mainclass()


response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>���̨��---�������</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=zjtz_class.asp>�������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=zjtz_class.asp?action=addclass>��ӷ���</a>"
response.write "</td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

  dim sqlbody,rsbody
  sqlbody="SELECT * from class order by id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connzj,1,1
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
  
     response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     response.write "<tr class=""title"">"
     response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ ��</strong></div></td>"
     response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("id")&"</div></td>"
        response.write " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("name")&"&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
if session("level")=0 or session("level")=9 then 	   
response.write "<a href=zjtz_class.asp?action=editclass&id="&rsbody("id")&">�༭</a>&nbsp;&nbsp;<a href=zjtz_class.asp?action=delclass&id="&rsbody("id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
	END IF    
response.write "&nbsp;</div></td></tr>"
	    RowCount=RowCount-1
    rsbody.movenext
    loop
end if 
     call showpage1(page,url,total,record,PgSz)
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub

response.write "</body></html>"


Call CloseConn
%>