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
'dim sqlfdbw,rsfdbw,title,record,pgsz,total,page,start,rowcount,xh,url,ii
'dim rsadd,sqladd,fdbwid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,fdbwwh,sql,rs,czjg
dim fdbwid,fdbwwh,sql,rs,sqlfdbw,rsfdbw,rsadd,sqladd,rsedit,sqledit
dim record,pgsz,total,page,start,rowcount,url,ii
dim czjg,id,rsdel,sqldel
fdbwid=Trim(Request("fdbwid"))
'fdbwwh=trim(request("fdbwwh"))	

dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>��Ϣ����ϵͳ�������¹���ҳ</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "<SCRIPT language=javascript>" & vbCrLf
dwt.out  "function checkadd(){" & vbCrLf
dwt.out  " if(document.form1.fdbw_whjl_body.value==''){" & vbCrLf
dwt.out  "      alert('ά�����ݲ���Ϊ�գ�');" & vbCrLf
dwt.out  "   document.form1.fdbw_whjl_body.focus();" & vbCrLf
dwt.out  "      return false;" & vbCrLf
dwt.out  "    }" & vbCrLf
dwt.out  "    }" & vbCrLf
dwt.out  "</SCRIPT>" & vbCrLf
dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if Request("action")="add" then call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del
if request("action")="" then call main 
sub add()
	fdbwwh=Connjg.Execute("SELECT wh FROM fdbw WHERE id="&fdbwid)(0)
'	sql="SELECT * from fdbw where id="&fdbwid
'	set rs=server.createobject("adodb.recordset")
'	rs.open sql,connjg,1,1
'	fdbwwh=rs("wh")
'	rs.close
	dwt.out "<br><br><br><form method='post' action='fdbw_whjl.asp' name='form1'  onsubmit='javascript:return checkadd();' >"
	dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out "<tr class='title'><td height='22' colspan='2'>"
	dwt.out "<div align='center'><strong>��ӷ�������  "&fdbwwh&"  ά����¼</strong></div></td>    </tr>"
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>ά��ԭ��</strong></td>"
	dwt.out "<td width='88%' class='tdbg'><input name='fdbw_whjl_whyy' type='text'></td>    </tr>   "
	
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ά��ʱ�䣺</strong></td> "
	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='fdbw_whjl_whsj' type='text' value="&date()&" onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out "</td></tr>"& vbCrLf
	
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ά�����ݣ�</strong></td> "
	dwt.out "<td><input name='fdbw_whjl_body' type='text'></td></tr>"
	
	   id=Trim(Request("fdbwid"))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from fdbw where id="&id
   rsedit.open sqledit,connjg,1,1

		dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='fdbw_tyqk' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("tyqk")=1 then dwt.out" selected"
	dwt.out">Ͷ��</option>"
	dwt.out"<option value='2'"
	if rsedit("tyqk")=2 then dwt.out" selected"
	dwt.out">�߱�����</option>"
	dwt.out"<option value='3'"
	if rsedit("tyqk")=3 then dwt.out" selected"
	dwt.out">��ȱ��</option>"
	dwt.out"<option value='4'"
	if rsedit("tyqk")=4 then dwt.out" selected"
	dwt.out">����ȡ��</option>"
    dwt.out"</select></td></tr>"
	       rsedit.close
       set rsedit=nothing

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd'> <input name='fdbw_whjl_fdbwid' type='hidden'  value='"&Trim(Request("fdbwid"))&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from fdbw_whjl" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("whyy")=Trim(Request("fdbw_whjl_whyy"))
      rsadd("body")=request("fdbw_whjl_body")
      rsadd("whsj")=Trim(request("fdbw_whjl_whsj"))
      rsadd("whjg")=request("fdbw_tyqk")
	  fdbwid=request("fdbw_whjl_fdbwid")
      rsadd("fdbwid")=trim(request("fdbw_whjl_fdbwid"))
      
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from fdbw where id="&Trim(request("fdbw_whjl_fdbwid"))
      rsedit.open sqledit,connjg,1,3
      	  rsedit("update")=now()
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	
	set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from fdbw where id="&ReplaceBadChar(trim(request("fdbw_whjl_fdbwid")))
      rsedit.open sqledit,connjg,1,3
	  rsedit("tyqk")=request("fdbw_tyqk")

	  rsedit.update
      rsedit.close
      set rsedit=nothing
	
	  dwt.savesl "��������ά����¼","�½�",Connjg.Execute("SELECT wh FROM fdbw WHERE id="&trim(request("fdbw_whjl_fdbwid"))&"")(0) 
	  dwt.out "<Script Language=Javascript>location.href='fdbw_whjl.asp?fdbwid="&fdbwid&"';</Script>"
end sub


sub saveedit()    
	  '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from fdbw_whjl where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connjg,1,3
      rsedit("whyy")=Trim(Request("fdbw_whjl_whyy"))
      rsedit("body")=request("fdbw_whjl_body")
      rsedit("whsj")=Trim(request("fdbw_whjl_whsj"))
	  rsedit("whjg")=request("fdbw_tyqk")
	  
	  
	  set rs=server.createobject("adodb.recordset")
      sql="select * from fdbw where id="&rsedit("fdbwid")
      rs.open sql,connjg,1,3
	  rs("tyqk")=request("fdbw_tyqk")
	  rs.update
      rs.close
      set rs=nothing

	  dwt.savesl "��������ά����¼","�༭",Connjg.Execute("SELECT wh FROM fdbw WHERE id="&rsedit("fdbwid")&"")(0) 

      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out "<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
 	sqledit="select * from fdbw_whjl where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
    dwt.savesl "��������ά����¼","ɾ��",Connjg.Execute("SELECT wh FROM fdbw WHERE id="&rsedit("fdbwid")&"")(0) 
	rsedit.close
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from fdbw_whjl where id="&id
  rsdel.open sqldel,connjg,1,3
  dwt.out "<Script Language=Javascript>history.back()</Script>"
set rsdel=nothing  

end sub


sub edit()
  sql="SELECT * from fdbw where id="&fdbwid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
fdbwwh=rs("wh")
rs.close
 id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from fdbw_whjl where id="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out "<br><br><br><form method='post' action='fdbw_whjl.asp' name='form1'  onsubmit='javascript:return checkadd();' >"
   dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>�༭��������  "&fdbwwh&"  ������¼</strong></div></td>    </tr>"
     
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out "<strong>ά��ԭ��</strong></td>"
	 dwt.out "<td width='88%' class='tdbg'><input name='fdbw_whjl_whyy' type='text' value="&rsedit("whyy")&"></td>    </tr>   "
	 
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ά��ʱ�䣺</strong></td> "
   dwt.out "<td width='88%' class='tdbg'>"
   dwt.out "<input name='fdbw_whjl_whsj' type='text' value="&rsedit("whsj")&" onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out "</td></tr>"& vbCrLf
   
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ά�����ݣ�</strong></td> "
	dwt.out "<td><input name='fdbw_whjl_body' type='text' value="&rsedit("body")&"></td></tr>"
		dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='fdbw_tyqk' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("whjg")=1 then dwt.out" selected"
	dwt.out">Ͷ��</option>"
	dwt.out"<option value='2'"
	if rsedit("whjg")=2 then dwt.out" selected"
	dwt.out">�߱�����</option>"
	dwt.out"<option value='3'"
	if rsedit("whjg")=3 then dwt.out" selected"
	dwt.out">��ȱ��</option>"
	dwt.out"<option value='4'"
	if rsedit("whjg")=4 then dwt.out" selected"
	dwt.out">����ȡ��</option>"
    dwt.out"</select></td></tr>"

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveedit'><input type='hidden' name='id' value='"&id&"'> <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"

end sub


sub main()
dim lb,brxx
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  "<tr class='topbg'>"& vbCrLf
dwt.out  "<td height='22' colspan='2' align='center'><strong>�������£�������¼</strong></td>"& vbCrLf
dwt.out  "</tr>"& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "<td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
dwt.out  "<td height='30'><a href=""fdbw.asp"">����������ҳ</a>&nbsp;|&nbsp;<a href=""fdbw.asp?action=add"">��ӷ�������</a>"
dwt.out  "</td>"& vbCrLf
dwt.out  "  </tr>"& vbCrLf
dwt.out  "</table>"& vbCrLf

sql="SELECT * from fdbw where id="&fdbwid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
if session("levelclass")=rs("sscj") or session("level")=0 then 
	dwt.out  "<a href='fdbw_whjl.asp?action=add&fdbwwh="&fdbwwh&"&fdbwid="&fdbwid&"'>��ӷ�������<font color='#ff0000'>"&rs("wh")&"</font>������¼</a>"
 else
    dwt.out  "&nbsp;"
 end if 
 fdbwwh=rs("wh")
 		select case rs("lb")
          case 1
             lb="һ"
          case 2 
        	lb="��"
        end select	 
		select case rs("brxx")
          case 1
             brxx="��"
          case 2 
        	brxx="��"
        end select	 

 dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
dwt.out  "<tr class=""title"">"  & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center""><strong>����</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>λ��</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������ʽ</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ʼʱ��</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>Ͷ��ʱ��</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ͣ��ʱ��</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ע</strong></div></td>" & vbCrLf
dwt.out  "    </tr>" & vbCrLf
                 dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&sscjh_d(rs("sscj"))&"</div></td>" & vbCrLf
                dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&gh(rs("ssgh"))&"&nbsp;</td>" & vbCrLf
                dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rs("wh")&"&nbsp;</td>" & vbCrLf
                dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("jz")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&lb&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&brxx&"&nbsp;</div></td>" & vbCrLf
	            dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("csdate")&"&nbsp;</div></td>" & vbCrLf
	            dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("date")&"&nbsp;</div></td>" & vbCrLf
	            dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("tydate")&"&nbsp;</div></td>" & vbCrLf
		        dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rs("bz")&"&nbsp;</td>" & vbCrLf

 dwt.out  " </tr></table>"
rs.close
set rs=nothing


dwt.out  "<div align='center'>ά����¼</div>"
sqlfdbw="SELECT * from fdbw_whjl where fdbwid="&fdbwid&" ORDER BY id DESC"
set rsfdbw=server.createobject("adodb.recordset")
rsfdbw.open sqlfdbw,connjg,1,1
if rsfdbw.eof and rsfdbw.bof then 
dwt.out  "<p align='center'>δ��ӷ�������<font color='#ff0000'>"&fdbwwh&"</font>������¼</p>" 
else
dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out  "<tr class=""title"">" 
dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
dwt.out  "      <td width=""40%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ά��ԭ��</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ά��ʱ��</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ά������</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ά�����</strong></div></td>"
dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"

dwt.out  "    </tr>"
           record=rsfdbw.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsfdbw.PageSize = Cint(PgSz) 
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
           rsfdbw.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsfdbw.PageSize
           do while not rsfdbw.eof and rowcount>0
      		
			dim tyqk
				select case rsfdbw("whjg")
			  case 1
				 tyqk="<span style='color:#006600'>Ͷ��</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>�߱�����</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>��ȱ��</span>"
			  case 4 
				tyqk="����ȡ��"
			end select	 
                 dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
          dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rsfdbw("id")&"</div></td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rsfdbw("whyy")&"&nbsp;</td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("whsj")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("body")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&tyqk&"&nbsp;</div></td>"
                dwt.out  "<td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				sql="SELECT * from fdbw where id="&fdbwid
                set rs=server.createobject("adodb.recordset")
                rs.open sql,connjg,1,1
				call editdel(rsfdbw("id"),rs("sscj"),"fdbw_whjl.asp?action=edit&fdbwid="&fdbwid&"&id=","fdbw_whjl.asp?action=del&id=")
				rs.close
                set rs=nothing

                dwt.out  "</div></td></tr>"
                 RowCount=RowCount-1
          rsfdbw.movenext
          loop
        dwt.out  "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rsfdbw.close
       set rsfdbw=nothing
        connjg.close
        set connjg=nothing

end sub
dwt.out  "</body></html>"
Call Closeconn
%>