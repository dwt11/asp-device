<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<%
'dim sqllsda,rslsda,title,record,pgsz,total,page,start,rowcount,xh,url,ii
'dim rsadd,sqladd,lsdaid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,lsdawh,sql,rs,czjg
dim lsdaid,lsdawh,sql,rs,sqllsda,rslsda,rsadd,sqladd,rsedit,sqledit
dim record,pgsz,total,page,start,rowcount,xh,url,ii
dim czjg,id,rsdel,sqldel
lsdaid=Trim(Request("lsdaid"))
'lsdawh=trim(request("lsdawh"))	

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title> ��Ϣ����ϵͳ������������ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.lsda_czjl_czyy.value==''){" & vbCrLf
dwt.out "      alert('����ԭ����Ϊ�գ�');" & vbCrLf
dwt.out "   document.form1.lsda_czjl_czyy.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if Request("action")="add" then call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del
if request("action")="" then call main 

sub add()
'// sql="SELECT * from lsda where lsdaid="&lsdaid
'//set rs=server.createobject("adodb.recordset")
'//rs.open sql,connjg,1,1
'//	lsdawh=rs("wh")
'//rs.close
'//set rs=nothing
  dwt.out"<br><br><br><form method='post' action='lsda_czjl.asp' name='form1' onsubmit='javascript:return checkadd();' >"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>������� "&connjg.ExeCute("SELECT wh FROM lsda where lsdaid="&lsdaid)(0)&"  ������¼</strong></div></td>    </tr>"
	 
	 	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td> "
	dwt.out"<td><select name='lsda_czjl_czyy' size='1'>"
	dwt.out"<option value='true'>ԭ��</option>"
    dwt.out"<option value='false'>����ԭ��</option>"
    dwt.out"</select></td></tr>"

	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>����˵����</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='lsda_czjl_czinfo' type='text'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='lsda_czjl_czsj' type='text' value="&date()&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, lsda_czjl_czsj, ' yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td> "
	dwt.out"<td><select name='lsda_czjl_czjg' size='1'>"
	dwt.out"<option value='1'>Ͷ��</option>"
    dwt.out"<option value='0'>��·</option>"
    dwt.out"</select></td></tr>"

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'> <input name='lsda_czjl_lsdaid' type='hidden'  value='"&Trim(Request("lsdaid"))&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
		dwt.out"�����Ͷ��,����ԭ��ɲ�ѡ��"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from lsda_czjl" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("czyy")=Trim(Request("lsda_czjl_czyy"))
      rsadd("czinfo")=Trim(Request("lsda_czjl_czinfo"))
      rsadd("czjg")=request("lsda_czjl_czjg")
      rsadd("czsj")=Trim(request("lsda_czjl_czsj"))
      rsadd("lsdaid")=trim(request("lsda_czjl_lsdaid"))
	  lsdaid=request("lsda_czjl_lsdaid")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  '����������������Ӧλ�ŵ�Ͷ��״��
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&Trim(request("lsda_czjl_lsdaid"))
      rsedit.open sqledit,connjg,1,3
      rsedit("tyzk")=request("lsda_czjl_czjg")
      rsedit("czyy")=request("lsda_czjl_czyy")
      	  rsedit("update")=now()
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  dwt.savesl "��������������¼","���",Connjg.Execute("SELECT wh FROM lsda WHERE lsdaid="&trim(request("lsda_czjl_lsdaid"))&"")(0) 
	  
	  dwt.out"<Script Language=Javascript>location.href='lsda_czjl.asp?lsdaid="&lsdaid&"';</Script>"
end sub


sub saveedit()    
	  '����
     set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda_czjl where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connjg,1,3
      rsedit("czyy")=Trim(Request("lsda_czjl_czyy"))
      rsedit("czinfo")=Trim(Request("lsda_czjl_czinfo"))
      rsedit("czjg")=request("lsda_czjl_czjg")
      rsedit("czsj")=Trim(request("lsda_czjl_czsj"))
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  	  '����������������Ӧλ�ŵ�Ͷ��״��
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&Trim(request("lsda_czjl_lsdaid"))
      rsedit.open sqledit,connjg,1,3
      rsedit("tyzk")=request("lsda_czjl_czjg")
      rsedit("czyy")=request("lsda_czjl_czyy")
	  rsedit("update")=now()
      rsedit.update
      rsedit.close
      set rsedit=nothing

	  dwt.savesl "��������������¼","�༭",Connjg.Execute("SELECT wh FROM lsda WHERE lsdaid="&trim(request("lsda_czjl_lsdaid"))&"")(0) 
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
 	sqledit="select * from lsda_czjl where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
    dwt.savesl "��������������¼","ɾ��",Connjg.Execute("SELECT wh FROM lsda WHERE lsdaid="&rsedit("lsdaid")&"")(0) 
	rsedit.close
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from lsda_czjl where id="&id
  rsdel.open sqldel,connjg,1,3
  dwt.out"<Script Language=Javascript>history.back()</Script>"
set rsdel=nothing  

end sub


sub edit()
' sql="SELECT * from lsda where lsdaid="&lsdaid
'set rs=server.createobject("adodb.recordset")
'rs.open sql,connjg,1,1
'	lsdawh=rs("wh")
'rs.close
'set rs=nothing

   id=Trim(request("id"))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from lsda_czjl where id="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><br><form method='post' action='lsda_czjl.asp' name='form1'  onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�༭����  "&connjg.ExeCute("SELECT wh FROM lsda where lsdaid="&lsdaid)(0)&"   ������¼</strong></div></td>    </tr>"
     
	 	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td> "
	dwt.out"<td><select name='lsda_czjl_czyy' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("czjg")=true then dwt.out"selected"
	dwt.out">ԭ��</option>"
    dwt.out"<option value='0'"
	if rsedit("czjg")=false then dwt.out"selected"
	dwt.out">����ԭ��</option>"
    dwt.out"</select></td></tr>"

	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>����˵����</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='lsda_czjl_czinfo' type='text' value="&rsedit("czinfo")&"></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='lsda_czjl_czsj' type='text' value="&rsedit("czsj")&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, lsda_czjl_czsj, ' yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
   
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td> "
	dwt.out"<td><select name='lsda_czjl_czjg' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("czjg")=1 then dwt.out"selected"
	dwt.out">Ͷ��</option>"
    dwt.out"<option value='0'"
	if rsedit("czjg")=0 then dwt.out"selected"
	dwt.out">��·</option>"
    dwt.out"</select></td></tr>"

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'><input name='lsda_czjl_lsdaid' type='hidden'  value='"&Trim(Request("lsdaid"))&"'>   <input type='hidden' name='id' value='"&id&"'> <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table>		</form>"
dwt.out"�����Ͷ��,����ԭ��ɲ�ѡ��"
end sub


sub main()
' sql="SELECT * from lsda where lsdaid="&lsdaid
'set rs=server.createobject("adodb.recordset")
'rs.open sql,connjg,1,1
'	lsdawh=rs("wh")
'rs.close
'set rs=nothing

dwt.out "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out "<tr class='topbg'>"& vbCrLf
dwt.out "<td height='22' colspan='2' align='center'><strong>����������������¼</strong></td>"& vbCrLf
dwt.out "</tr>"& vbCrLf
dwt.out "<tr class='tdbg'>"& vbCrLf
dwt.out "<td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
dwt.out "<td height='30'><a href=""lsda.asp"">����������ҳ</a>&nbsp;|&nbsp;<a href=""lsda.asp?action=add"">�����������</a>"
dwt.out "</td>"& vbCrLf
dwt.out "  </tr>"& vbCrLf
dwt.out "</table>"& vbCrLf

				
sql="SELECT * from lsda where lsdaid="&lsdaid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
if session("levelclass")=rs("sscj") or session("levelclass")=10 then 
	dwt.out "<a href='lsda_czjl.asp?action=add&lsdaid="&lsdaid&"'>�������<font color='#ff0000'>"&connjg.ExeCute("SELECT wh FROM lsda where lsdaid="&lsdaid)(0)&"</font>������¼</a>"
 else
    dwt.out "&nbsp;"
 end if 
 rs.close
set rs=nothing



sqllsda="SELECT * from lsda_czjl where lsdaid="&lsdaid&" ORDER BY id DESC"
set rslsda=server.createobject("adodb.recordset")
rslsda.open sqllsda,connjg,1,1
if rslsda.eof and rslsda.bof then 
dwt.out "<p align='center'>δ�������<font color='#ff0000'>"&lsdawh&"</font>������¼</p>" 
else
dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out "<tr class=""title"">" 
dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>����λ��</strong></div></td>"
dwt.out "      <td width=""40%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ԭ��</strong></div></td>"
dwt.out "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ʱ��</strong></div></td>"
dwt.out "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�������</strong></div></td>"
dwt.out "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"

dwt.out "    </tr>"
           record=rslsda.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rslsda.PageSize = Cint(PgSz) 
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
           rslsda.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rslsda.PageSize
           do while not rslsda.eof and rowcount>0
		'xh=xh+1
                 dwt.out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rslsda("id")&"</div></td>"
                dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%"">"&connjg.ExeCute("SELECT wh FROM lsda where lsdaid="&lsdaid)(0)&"</td>"
                dwt.out "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rslsda("czinfo")&"&nbsp;</td>"
                dwt.out "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rslsda("czsj")&"&nbsp;</div></td>"
        select case rslsda("czjg")
          case 0
            czjg="��·"
           if rslsda("czyy") then
		    czjg="<font color='#ff0000'>"&czjg&"</font>"
		   else
		    czjg="<font color='#0000ff'>"&czjg&"</font>"
		   end if 	
		  case 1 
        	czjg="Ͷ��"
        end select	 
		dwt.out "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&czjg&"&nbsp;</div></td>"
                dwt.out "<td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				sql="SELECT * from lsda where lsdaid="&lsdaid
                set rs=server.createobject("adodb.recordset")
                rs.open sql,connjg,1,1
				call editdel(rslsda("id"),rs("sscj"),"lsda_czjl.asp?action=edit&lsdaid="&lsdaid&"&id=","lsda_czjl.asp?action=del&id=")
				rs.close
                set rs=nothing

                dwt.out "</div></td></tr>"
                 RowCount=RowCount-1
          rslsda.movenext
          loop
        dwt.out "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rslsda.close
       set rslsda=nothing
        connjg.close
        set connjg=nothing

end sub
dwt.out "</body></html>"
Call Closeconn
%>