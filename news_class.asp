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
dim sqluser,rsuser,sqlcj,rscj
url="news_class.asp"
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
Dwt.out "<html>"
Dwt.out "<head>"
Dwt.out "<title>����ҳ�������</title>"
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function CheckAdd(){" & vbCrLf
 Dwt.out " if(document.form1.class_name.value==''){" & vbCrLf
Dwt.out "      alert('���Ʋ���Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form1.class_name.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out "</head>"
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"


sub add()
   '����
   Dwt.out"<form method='post' action='news_class.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>��������ҳ����</strong></Div></td>    </tr>"
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 Dwt.out"<strong>��������</strong></td>"
	 Dwt.out"<td width='80%' class='tdbg'><input name='class_name' type='text'></td>    </tr>   "
   	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�Ƿ���ҳ��ʾ��</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'>"
	 Dwt.out"<select name=isre><option value=true>��ʾ</option><option value=false>����ʾ</option></select></td>    </tr> "
   	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�Ƿ���Իظ���</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'>"
	 Dwt.out"<select name=isindex><option value=true>��</option><option value=false>��</option></select></td>    </tr> "
 Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='news_class.asp';"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xzgl_news_class" 
      rsadd.open sqladd,connxzgl,1,3
      rsadd.addnew
      rsadd("class_name")=ReplaceBadChar(Trim(Request("class_name")))
      rsadd("isindex")=Request("isindex")
      rsadd("isre")=Request("isre")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
	
	  
end sub

sub main()
     	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>����ҳ�������</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
'�û�������ҳ
	Dwt.out "<Div class='x-toolbar'><Div align=left><a href='news_class.asp?action=add'>��ӷ���</a></Div></Div>" & vbCrLf
 		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf

	  Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      Dwt.out "<tr  class=""x-grid-header"">" 
      Dwt.out "     <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center""><strong>���</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>������</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>��ַ</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>�Ƿ���ҳ��ʾ</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>�Ƿ���Իظ�</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>����</strong></Div></td>"
      Dwt.out "    </tr>"
      sqlbody="SELECT * from xzgl_news_class "
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,connxzgl,1,1
      if rsbody.eof and rsbody.bof then 
           Dwt.out "<p align=""center"">��������</p>" 
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
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center"">"&xh_id&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("class_name")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">news.asp?classid="&rsbody("id")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("isindex")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("isre")&"</Div></td>"
                  Dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><a href='news_class.asp?action=edit&ID="&rsbody("id")&"'>�༭</a>&nbsp;"
				 Dwt.out "  <a href='news_class.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('ȷ��Ҫɾ��������������ɾ��������صĴ˷������ݽ�������ʾ');"">ɾ��</a></Div></td>"
                 Dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
		Dwt.out "</table>"& vbCrLf
		call showpage1(page,url,total,record,PgSz)
		Dwt.out "</Div>"& vbCrLf
       end if
 	Dwt.out "</Div>"  
      rsbody.close
       set rsbody=nothing
       
end sub

sub edit()
     '�༭
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from xzgl_news_class where id="&id
   rsedit.open sqledit,connxzgl,1,1

   Dwt.out"<form method='post' action='news_class.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�༭�������</strong></Div></td>    </tr>"
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����������ƣ�</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'>"
	 Dwt.out"<input type='text' name='class_name' value="&rsedit("class_name")&"></td>    </tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�Ƿ���ҳ��ʾ��</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'>"
	 Dwt.out"<select name=isindex><option value=true"
	 if rsedit("isindex") then dwt.out " selected"
	 dwt.out ">��ʾ</option><option value=false"
	 if rsedit("isindex")=false then dwt.out " selected"
	 dwt.out">����ʾ</option></select></td>    </tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�Ƿ���Իظ���</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'>"
	 Dwt.out"<select name=isre><option value=true"
	 if rsedit("isre") then dwt.out " selected"
	 dwt.out ">��</option><option value=false"
	 if rsedit("isre")=false then dwt.out " selected"
	 dwt.out">��</option></select></td>    </tr> "
    Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from xzgl_news_class where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connxzgl,1,3
rsedit("class_name")=ReplaceBadChar(Trim(Request("class_name")))
rsedit("isindex")=Request("isindex")
rsedit("isre")=Request("isre")
rsedit.update
rsedit.close
	
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from xzgl_news_class where id="&id
rsdel.open sqldel,connxzgl,1,3
Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub

Dwt.out "</body></html>"

Call CloseConn
%>