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
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>�豸�������ݹ���ҳ</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out "  if(document.form.sbjxnr_name.value==''){" & vbCrLf
Dwt.out "      alert('����δ��д��');" & vbCrLf
Dwt.out "  document.form.sbjxnr_name.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
select case Request("action")
  case ""
      call mainclass'��ҳ����ʾ��
  case "mainclass"
      call mainclass'��ҳ����ʾ��
  case "main"
      call main'��
  case "addclass"
      call addclass '���Ӹ�
  case "saveaddclass"
      call saveaddclass    '���游
  case "editclass"
      call editclass '�༭��
  case "saveeditclass"
      call saveeditclass '�༭���游
  case "delclass"
      call delclass  'ɾ������Ϣ
end select	  

sub addclass()'���
   Dwt.out"<form method='post' action='sb_jxnr_class.asp' name='form' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�豸�����������</strong></Div></td>    </tr>"
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ƣ� </strong></td>"      
    Dwt.out"<td width='88%' class='tdbg'>"
       Dwt.out"<input name='sbjxnr_name' type='text'></td></tr>"& vbCrLf

    dim rs,sql,rsz,sqlz
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�����ϼ��� </strong></td>"      
    Dwt.out"<td width='88%' class='tdbg'>"

Dwt.out "<select name='sb_jxnr_class' size='1' >"
Dwt.out "  <option selected value='0'>ѡ��һ��</option>"
	sql="SELECT * from sbjxnr where sbjxnr_zclass=0 "& vbCrLf
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    do while not rs.eof
       	Dwt.out"<option value='"&rs("sbjxnr_id")&"'>"&rs("sbjxnr_name")&"</option>"& vbCrLf
		rs.movenext
	loop
	rs.close
	set rs=nothing
	Dwt.out "</select>"
	
	
	
	
	
	 
		 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbjxnr_orderby' type='text'></td></tr>"& vbCrLf
   
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveaddclass'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
	message "ʲô����ѡ������һ��;"
end sub	

sub saveaddclass()    
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from sbjxnr" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
       'on error resume next
      if request("sb_jxnr_class")=0 then 
	     rsadd("sbjxnr_zclass")=0
      else
	     if request("sb_zclass")=0 then 
		    rsadd("sbjxnr_zclass")=ReplaceBadChar(request("sb_jxnr_class"))
		 else
		    rsadd("sbjxnr_zclass")=ReplaceBadChar(request("sb_zclass"))
		 end if 
      end if 
	  rsadd("sbjxnr_name")=ReplaceBadChar(request("sbjxnr_name"))
	  rsadd("sbjxnr_orderby")=0
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>location.href='sb_jxnr_class.asp?action=mainclass';</Script>"
end sub
 



sub saveeditclass()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbjxnr where sbjxnr_id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conn,1,3
      rsedit("sbjxnr_name")=ReplaceBadChar(Trim(Request("sbjxnr_name")))
	  	  	  rsedit("sbjxnr_orderby")=ReplaceBadChar(request("sbjxnr_orderby"))
		  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub




sub delclass()
dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from sbjxnr where sbjxnr_id="&request("id")
  rsdel.open sqldel,conn,1,3
  Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub








sub editclass()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from sbjxnr where sbjxnr_id="&id
   rsedit.open sqledit,conn,1,1
   Dwt.out"<form method='post' action='sb_jxnr_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�༭</strong></Div></td>    </tr>"
     
     Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ƣ� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbjxnr_name' type='text' value='"&rsedit("sbjxnr_name")&"'></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbjxnr_orderby' type='text' value='"&rsedit("sbjxnr_orderby")&"'></td></tr>"& vbCrLf

		Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveeditclass'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
	
       rsedit.close
       set rsedit=nothing
end sub

'�ж��Ƿ�����
function zclassor(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbjxnr where sbjxnr_zclass="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclassor=false 
  else
     zclassor=true
  end if
end function



Sub mainclass()
  	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�豸�������ݹ���</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
	Dwt.out "<Div class='x-toolbar'>" & vbCrLf
	Dwt.out "<Div align=left><a href=""sb_jxnr_class.asp?action=addclass"">���</a></Div>" & vbCrLf
	Dwt.out "</Div>"

  dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxnr where sbjxnr_zclass=0 order by  sbjxnr_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
  	 Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
     Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     Dwt.out "<tr class=""x-grid-header"">"
     Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�����ϼ�</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�� ��</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>ѡ ��</Div></td>"
     Dwt.out "    </tr>"
  
  do while not rsbody.eof 
	  dim xh,xh_id
		xh=xh+1
			if xh mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
        Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"</Div></td>"
        Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><b>"&rsbody("sbjxnr_name")&"</b>&nbsp;</Div></td>"
        Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">һ��</Div></td>"
        Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("sbjxnr_orderby")&"&nbsp;</Div></td>"
       Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
	   Dwt.out "<a href=sb_jxnr_class.asp?action=editclass&id="&rsbody("sbjxnr_id")&">�༭</a>&nbsp;&nbsp;<a href=sb_jxnr_class.asp?action=delclass&id="&rsbody("sbjxnr_id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
	   Dwt.out "</Div></td></tr>"
	    			'����
			sqlz="SELECT * from sbjxnr where sbjxnr_zclass="&rsbody("sbjxnr_id")&" order by  sbjxnr_orderby aSC"& vbCrLf
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conn,1,1
			if rsz.eof and rsz.bof then 
			else
				dim xhz
				xhz=0
				do while not rsz.eof
				
					xhz=xhz+1
					if xhz mod 2 =1 then 
					  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"-"&xhz&"</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""left"">"&rsz("sbjxnr_name")&"&nbsp;</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&conn.Execute("SELECT sbjxnr_name FROM sbjxnr WHERE  sbjxnr_id="&rsz("sbjxnr_zclass"))(0)&"-����</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsz("sbjxnr_orderby")&"&nbsp;</Div></td>"
				   Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
				   dwt.out "<a href=sb_jxnr_class.asp?action=editclass&id="&rsz("sbjxnr_id")&">�༭</a>&nbsp;&nbsp;<a href=sb_jxnr_class.asp?action=delclass&id="&rsz("sbjxnr_id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
				   Dwt.out "</Div></td></tr>"
				
				rsz.movenext
				loop
			end if 	
			rsz.close
			set rsz=nothing
		
    rsbody.movenext
    loop
     Dwt.out "</table></Div>"
end if 
  rsbody.close
  set rsbody=nothing
  'conn.close
  'set conn=nothing
  Dwt.out "</Div>"
end sub

Dwt.out "</body></html>"



Call CloseConn
%>