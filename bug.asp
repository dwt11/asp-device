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
dim sqlbug,rsbug,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
url="bug.asp"

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ�����ռ�ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function CheckAdd(){" & vbCrLf
 response.write " if(document.form1.bug_title.value==''){" & vbCrLf
response.write "      alert('���ⲻ��Ϊ�գ�');" & vbCrLf
response.write "   document.form1.bug_title.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.bug_body.value==''){" & vbCrLf
response.write "      alert('���ݲ���Ϊ�գ�');" & vbCrLf
response.write "  document.form1.bug_body.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>��Ϣ����ϵͳ�����ռ�ҳ</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='100%' height='30'><strong>��ҳΪ����ڣ���Ϣ����ϵͳ��ʹ���г��ֵĴ����Լ�δʵ�ֵĹ��ܵȷ��������ռ�ҳ��</strong></td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='100%' height='30'><strong><a href='bug.asp?action=add'>��Ӽ���</a></td>"& vbCrLf
response.write "  </tr>"& vbCrLf

response.write "</table>"& vbCrLf

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
   response.write"<form method='post' action='bug.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>��Ӽ���</strong></div></td>    </tr>"
		 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
 	 response.write"<strong>��&nbsp;&nbsp;�⣺</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='bug_title' type='text'></td>    </tr>   "
response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�����ˣ�</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='text' name='bug_user' disabled='true' value="&session("username")&"></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʱ&nbsp;&nbsp;�䣺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input type='text' name='bug_date' disabled='true' value="&now()&"></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ݣ� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=bug_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe>  <input type='hidden' name='bug_body' value=''>"
    response.write"</td></tr>  "   
    response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='bug.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from bug" 
      rsadd.open sqladd,connbug,1,3
      rsadd.addnew
                rsadd("bug_title")=ReplaceBadChar(Trim(Request("bug_title")))
  rsadd("bug_user")=session("username")
      rsadd("bug_body")=Trim(request("bug_body"))
      rsadd("bug_date")=now()

      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>window.alert('��Ӽ���ɹ�');location.href='bug.asp';</Script>"
end sub

sub edit()
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from bug where id="&id
   rsedit.open sqledit,connbug,1,1

   response.write"<form method='post' action='bug.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�� �� �� ��</strong></div></td>    </tr>"
      response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
   response.write"<strong>��&nbsp;&nbsp;�⣺</strong></td>"
   response.write"<td width='88%' class='tdbg'><input name='bug_title' type='text' value='"&rsedit("bug_title")&"'></td>    </tr>   "
response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�����ˣ�</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input name='bug_user' type='text'  disabled='true' value='"&rsedit("bug_user")&"'></td>    </tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʱ&nbsp;&nbsp;�䣺</strong></td> "
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<input name='bug_date' type='text' disabled='true' value='"&now()&"'></td>    </tr> "
	  
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ݣ� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	scontent=rsedit("bug_body")
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=bug_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe><textarea name='bug_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    response.write"</td></tr>  "   

	 
    response.write"<tr> <td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='usermanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from bug where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connbug,1,3
rsedit("bug_body")=Trim(request("bug_body"))
rsedit("bug_date")=now()
rsedit("bug_title")=ReplaceBadChar(Trim(Request("bug_title")))
rsedit.update
rsedit.close
	  response.write"<Script Language=Javascript>window.alert('�༭����ɹ�');location.href='bug.asp';</Script>"
	
end sub


sub main()
sqlbug="SELECT * from bug ORDER BY id DESC"
set rsbug=server.createobject("adodb.recordset")
rsbug.open sqlbug,connc,1,1
if rsbug.eof and rsbug.bof then 
response.write "<p align='center'>û���κμ���</p>" 
else
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ʱ��</strong></div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
response.write "    </tr>"
           record=rsbug.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsbug.PageSize = Cint(PgSz) 
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
           rsbug.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsbug.PageSize
           do while not rsbug.eof and rowcount>0
                 xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
                
                 response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=bug_view.asp?id="&rsbug("id")&" target=_blank>"&rsbug("bug_title")&"</a></td>"
				 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbug("bug_user")&"</div></td>"
                 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbug("bug_date")&"</div></td>"
 				 if session("level")=0 or session("username")=rsbug("bug_user") then
                 response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='bug.asp?action=edit&ID="&rsbug("id")&"'>�༭</a>&nbsp;"
   				  response.write "  <a href='bug.asp?action=del&ID="&rsbug("id")&"' onClick=""return confirm('ȷ��Ҫɾ���˼�����');"">ɾ��</a></div></td>"
                 else
				   response.write " <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
				 end if 
				 response.write "    </tr>"
                 RowCount=RowCount-1
          rsbug.movenext
          loop
           call showpage1(page,url,total,record,PgSz)
   end if
       rsbug.close
       set rsbug=nothing
        conn.close
        set conn=nothing
        response.write "</table>"
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from bug where id="&id
rsdel.open sqldel,connbug,1,3
response.write"<Script Language=Javascript>window.alert('ɾ������ɹ�');location.href='bug.asp';</Script>"
'rsdel.close
set rsdel=nothing  

end sub







response.write "</body></html>"

Call CloseConn
%>