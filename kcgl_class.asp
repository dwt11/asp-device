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
dim rs,sql,title
url="kcgl_class.asp?type="&request("type")
'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
if request("type")=1 then title="����"
if request("type")=2 then title="����"
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>"&title&"�������ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out "  if(document.form1.kcgl_name.value==''){" & vbCrLf
dwt.out "      alert('��������δ��д��');" & vbCrLf
dwt.out "  document.form1.kcgl_name.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if Request("action")="" then call main '��ҳ����ʾ�ӷ���
if Request("action")="add" then call add '�ӷ������
if Request("action")="saveadd" then call saveadd '�ӷ�����ӱ���
if Request("action")="edit" then call edit   '�༭�ӷ���
if Request("action")="saveedit" then call saveedit  '�༭�����ӷ���
if Request("action")="del" then call del     'ɾ���ӷ�����Ϣ


sub add()
   dwt.out"<form method='post' action='kcgl_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>���̨�ˣ�"&title&"�������</strong></div></td>    </tr>"
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
    dwt.out"<input name='kcgl_name' type='text'></td></tr>"& vbCrLf
  
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ࣺ </strong></td>"      
    dwt.out"<td><select name='kcgl_class' size='1'>"
	dwt.out"<option selected value='0'>ѡ��һ������</option>"
	dim sqlclass,rsclass
	 sqlclass="SELECT * from class where dclass=0 and type="&request("type")
    set rsclass=server.createobject("adodb.recordset")
    rsclass.open sqlclass,connkc,1,1
    if rsclass.eof and rsclass.bof then 
       dwt.out "�޷���" 
    else
	   do while not rsclass.eof
         dwt.out"<option value='"&rsclass("id")&"'>"&rsclass("name")&"</option>"
	   rsclass.movenext
	   loop
    end if
    dwt.out"</select></td></tr>"
	rsclass.close
	  
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
    dwt.out"<input name='kcgl_orderby' type='text'></td></tr>"& vbCrLf
  
	   
	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>   <input type='hidden' name='type' value='"&request("type")&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	


sub saveadd()    
	  '���浽�Դ����
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from class" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
       on error resume next
      rsadd("name")=request("kcgl_name")
      rsadd("type")=request("type")
      rsadd("dclass")=request("kcgl_class")
      rsadd("orderby")=request("kcgl_orderby")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub




sub saveedit()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connkc,1,3
      rsedit("name")=request("kcgl_name")
      rsedit("dclass")=request("kcgl_class")
      rsedit("orderby")=request("kcgl_orderby")
      rsedit("type")=request("type")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from class where id="&request("id")
  rsdel.open sqldel,connkc,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub
'sub delclass()
'dim rsdel,sqldel
'  set rsdel=server.createobject("adodb.recordset")
'  sqldel="delete * from class where id="&request("id")
'  rsdel.open sqldel,connkc,1,3
'  dwt.out"<Script Language=Javascript>history.go(-1)<Script>"
'set rsdel=nothing  
'
'end sub


sub edit()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from class where id="&id
   rsedit.open sqledit,connkc,1,1
   dwt.out"<form method='post' action='kcgl_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�༭���"&title&"��������</strong></div></td>    </tr>"
  	
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
    dwt.out"<input name='kcgl_name' type='text' value='"&rsedit("name")&"'></td></tr>"& vbCrLf
  
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ࣺ </strong></td>"      
    dwt.out"<td><select name='kcgl_class' size='1'>"
	dwt.out"<option selected value='0'>ѡ��һ������</option>"
	dim sqlclass,rsclass
	 sqlclass="SELECT * from class where dclass=0"
    set rsclass=server.createobject("adodb.recordset")
    rsclass.open sqlclass,connkc,1,1
    if rsclass.eof and rsclass.bof then 
       dwt.out "�޷���" 
    else
	   do while not rsclass.eof
         dwt.out"<option value='"&rsclass("id")&"'"
		 if rsclass("id")=rsedit("dclass") then dwt.out " selected"
		 dwt.out">"&rsclass("name")&"</option>"
	   rsclass.movenext
	   loop
    end if
    dwt.out"</select></td></tr>"
	rsclass.close
	  
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
    dwt.out"<input name='kcgl_orderby' type='text' value='"&rsedit("orderby")&"'></td></tr>"& vbCrLf
  
   

		dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>  <input type='hidden' name='type' value='"&request("type")&"'>   <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
       rsedit.close
       set rsedit=nothing
end sub



Sub main()
     
  	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>������"&title&"�������</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	dwt.out "<div class='x-toolbar'>" & vbCrLf
	dwt.out "<div align=left><a href=""kcgl_class.asp?action=add&type="&request("type")&""">��ӷ���</a></div>" & vbCrLf
	dwt.out "</div>"

  
  
  dim sqlbody,rsbody,xh
  sqlbody="SELECT * from class where type="&request("type")&" order by dclass,orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     dwt.out "<p align=""center"">��������</p>" 
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
	dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
	dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
	dwt.out "<tr class=""x-grid-header"">"
     dwt.out "<td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��������</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��������</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>ѡ ��</div></td>"
     dwt.out "    </tr>"
  
  do while not rsbody.eof and rowcount>0
		dim xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
		if xh_id mod 2 =1 then 
		  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
		else
		  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
		end if 
        dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh&"</div></td>"
        dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("name")&"&nbsp;</div></td>"
        dim classname
		if rsbody("dclass")=0 then 
		  classname="һ��"
		else 
		  classname=dclass(rsbody("dclass"))
		 end if 
		dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&classname&"&nbsp;</div></td>"
		dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("orderby")&"&nbsp;</div></td>"
       
	   dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
	   
	   dwt.out "<a href=kcgl_class.asp?type="&rsbody("type")&"&action=edit&id="&rsbody("id")&">�༭</a>&nbsp;&nbsp;<a href=kcgl_class.asp?action=del&id="&rsbody("id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
	   dwt.out "</div></td></tr>"
	    RowCount=RowCount-1
    rsbody.movenext
    loop
	dwt.out "</table>"
     call showpage1(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub


dwt.out "</body></html>"
function dclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from class where id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     dwt.out "���ִ���" 
  else
     dclass=rsbody("name")
  end if
end function
Call CloseConn
%>