<%@language=vbscript codepage=936 %>
<%
'Option Explicit
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
'���������ּ��ͱ���Ƿ���
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh")) 

dim sqllsda,rslsda,title,record,pgsz,total,page,start,rowcount,xh,url,ii,zxzz
dim rsadd,sqladd,lsdaid,rsedit,sqledit,scontent,rsdel,sqldel,tyzk,id

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>��Ϣ����ϵͳ������������ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.lsda_sscj.value==''){" & vbCrLf
dwt.out "      alert('��ѡ���������䣡');" & vbCrLf
dwt.out "   document.form1.lsda_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form1.lsda_wh.value==''){" & vbCrLf
dwt.out "      alert('λ�Ų���Ϊ�գ�');" & vbCrLf
dwt.out "   document.form1.lsda_wh.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
url="lsda.asp"
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
  case "deltrue"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call deltrue
  case "hy"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call hy
  case "savedel"
     call savedel
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call mainbody
  case "main"
    if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	


sub add()
dim rscj,sqlcj
   dwt.out"<br><br><br><form method='post' action='lsda.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�����������</strong></div></td>    </tr>"
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='lsda_sscj' size='1'>"
    dwt.out"<option >��ѡ����������</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
     dwt.out"<input name='lsda_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
      dwt.out"<input name='lsda_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf

 end if 

	 	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'>"
	 	dwt.out"<select name='lsda_gh' size='1' >"
     call formgh (0,session("levelclass"))
    dwt.out"</select> ��û����Ҫ�Ĺ���װ��,����ϵ�����������Ӧװ�ù�������</td></tr>"

	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='lsda_wh' type='text'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_yt' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>һ��Ԫ�����ƣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_ycjname'></td></tr> "
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_cldw'></td></tr>  "   
   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_clfw'></td></tr>  "   
   	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵL��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_lsl'></td></tr>  "   
   	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵH��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_lsh'></td></tr>  "   
   
	
    dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ִ��װ�ã�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_zxzz'></td></tr>  "   
    dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ִ��װ�ã�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"   
	 	  dwt.out"<select name='lsda_fj' size='1' >"
       dwt.out"<option value='0'>��ѡ��ּ�</option>"
		  dwt.out"<option value='1'>����</option>"
    dwt.out"<option value='2'>���</option>"
	dwt.out"<option value='3'>��</option>"
    dwt.out"</select></td></tr>"

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_bz'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from lsda" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("lsda_sscj"))
      rsadd("wh")=request("lsda_wh")
      rsadd("ssgh")=request("lsda_gh") '29���¼�
      rsadd("yt")=Trim(request("lsda_yt"))
      rsadd("ycjname")=request("lsda_ycjname")
      rsadd("cldw")=request("lsda_cldw")
      rsadd("clfw")=request("lsda_clfw")
      rsadd("lsl")=request("lsda_lsl")
      rsadd("lsh")=request("lsda_lsh")
      rsadd("zxzz")=request("lsda_zxzz")
	  rsadd("fj")=request("lsda_fj")   '29���¼�
	  'rsadd("update")=now()
      rsadd("bz")=request("lsda_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.savesl "��������","���",request("lsda_wh")
	  'dwt.out"<Script Language=Javascript>location.href='lsda.asp';<Script>"
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub deltrue()
  lsdaid=request("id")
 	sqledit="select * from lsda where lsdaID="&lsdaid
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
	dwt.savesl "��������","ɾ��",rsedit("wh")
	rsedit.close

  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from lsda where lsdaid="&lsdaid
  rsdel.open sqldel,connjg,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub

sub saveedit()    
	  '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
      rsedit("sscj")=Trim(Request("lsda_sscj"))
      rsedit("ssgh")=request("lsda_gh")   '29���¼�
      rsedit("wh")=request("lsda_wh")
      rsedit("yt")=Trim(request("lsda_yt"))
      rsedit("ycjname")=request("lsda_ycjname")
      rsedit("cldw")=request("lsda_cldw")
      rsedit("clfw")=request("lsda_clfw")
      rsedit("lsl")=request("lsda_lsl")
      rsedit("lsh")=request("lsda_lsh")
      rsedit("zxzz")=request("lsda_zxzz")
      rsedit("bz")=request("lsda_bz")
      rsedit("fj")=request("lsda_fj") '29���¼�
	  'rsedit("ssgh")=request("ssgh")
	  rsedit("update")=now()
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.savesl "��������","�༭",request("lsda_wh")
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()

   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from lsda where lsdaid="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><br><form method='post' action='lsda.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>ɾ����������</strong></div></td>    </tr>"
     
     dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'><input name='lsda_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out"<input name='lsda_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf
     
	  '29���¼�
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'>"
	 	dwt.out"<select name='lsda_gh' size='1'  disabled='disabled' >"
     call formgh (rsedit("ssgh"),rsedit("sscj"))
    dwt.out"</select> </td></tr>"


	 
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='lsda_wh' type='text'  disabled='disabled'  value='"&rsedit("wh")&"'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_yt'  disabled='disabled'   value='"&rsedit("yt")&"'></td></tr> "
	 
  
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��׼����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='deldate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"</td></tr>  "   
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��׼�������쵼</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='delr' ></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='savedel'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
       rsedit.close
       set rsedit=nothing
	






set rsdel=nothing  

end sub



SUB savedel()
  lsdaid=request("id")
 	sqledit="select * from lsda where lsdaID="&lsdaid
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
	dwt.savesl "��������","ɾ����ȡ����",rsedit("wh")
	rsedit.close

      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
      rsedit("isdel")=true
      rsedit("deldate")=request("deldate")
      rsedit("delr")=request("delr")
      rsedit.update
      rsedit.close
      set rsedit=nothing
  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
END SUB
SUB hy()
  lsdaid=request("id")
 	sqledit="select * from lsda where lsdaID="&lsdaid
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
	dwt.savesl "��������","��ԭ",rsedit("wh")
	rsedit.close

      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
      rsedit("isdel")=false
      rsedit.update
      rsedit.close
      set rsedit=nothing
  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
END SUB

sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from lsda where lsdaid="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><br><form method='post' action='lsda.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�༭��������</strong></div></td>    </tr>"
     
     dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'><input name='lsda_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out"<input name='lsda_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf
     
	  '29���¼�
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'>"
	 	dwt.out"<select name='lsda_gh' size='1' >"
     call formgh (rsedit("ssgh"),rsedit("sscj"))
    dwt.out"</select> ��û����Ҫ�Ĺ���װ��,����ϵ�����������Ӧװ�ù�������</td></tr>"


	 
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='lsda_wh' type='text' value='"&rsedit("wh")&"'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_yt'  value='"&rsedit("yt")&"'></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>һ��Ԫ�����ƣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_ycjname' value='"&rsedit("ycjname")&"'></td></tr> "
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_cldw' value='"&rsedit("cldw")&"'></td></tr>  "   
   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_clfw' value='"&rsedit("clfw")&"'></td></tr>  "   
   	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵL��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_lsl' value='"&rsedit("lsl")&"'></td></tr>  "   
   	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵH��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_lsh' value='"&rsedit("lsh")&"'></td></tr>  "   
   
	
    dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ִ��װ�ã�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_zxzz' value='"&rsedit("zxzz")&"'></td></tr>  "   
    
	 '29���¼�
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ּ���</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
	dwt.out"<select name='lsda_fj' size='1' >"
       dwt.out"<option value='0'"
	  if rsedit("fj")=0 then dwt.out " selected" 
	      dwt.out">��ѡ��ּ�</option>"
		  dwt.out"<option value='1' "
	   if rsedit("fj")=1 then dwt.out "selected"
	 dwt.out">����</option>"
    dwt.out"<option value='2'"
	if rsedit("fj")=2 then dwt.out "selected"
    dwt.out" >���</option>"
	dwt.out"<option value='3' "
	   if rsedit("fj")=3 then dwt.out "selected"
	 dwt.out">��</option>"
    dwt.out"</select></td></tr>"
  
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='lsda_bz' value='"&rsedit("bz")&"'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub


sub main()

dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>������������ҳ</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
dwt.out "<div class='x-toolbar'><div align=left><strong>��������</strong>"& vbCrLf
dwt.out "<a href=""lsda.asp?action=main"">����������ҳ</a>&nbsp;|&nbsp;<a href=""lsda.asp?action=add"">�����������</a>&nbsp;|&nbsp;<a href=""tocsv.asp?action=lsdamain&titlename=��������"" target=""_blank"">���������������̨�˵�Excel�ĵ�</a>  ����Խ�����Խ��Ҫ"& vbCrLf
dwt.out "  </div>"& vbCrLf
dwt.out "</div>"& vbCrLf
call search()
dwt.out "<br/><br/><br/>"



dwt.out "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out "<tr class=""title"">" 
dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" rowspan='2'><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" rowspan='2'><div align=""center""><strong>Ͷ������</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" rowspan='2'><div align=""center""><strong>��������</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" rowspan='2'><div align=""center""><strong>Ͷ����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" colspan='2'><div align=""center""><strong>δͶ������</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" colspan='3'><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" colspan='3'><div align=""center""><strong>���</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" colspan='3'><div align=""center""><strong>��</strong></div></td>"
dwt.out "    </tr>"
dwt.out "<tr class=""title"">" 
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong></strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>Ͷ��</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>δͶ��</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>Ͷ��</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>δͶ��</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>Ͷ��</strong></div></td>"
dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>δͶ��</strong></div></td>"

dwt.out "    </tr>"





    dim sqlcj,rscj
    sqlcj="SELECT * from levelname where levelclass=1 and levelid<4 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
		dwt.out "<tr class=""tdbg"" >" 
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><a href=lsda.asp?search=sscjs&sscj="&rscj("levelid")&">"&rscj("levelname")&"</a></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and not isdel and sscj="&rscj("levelid")&"")(0)&"</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and sscj="&rscj("levelid")&"")(0) &"</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=1 and sscj="&rscj("levelid")&"")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE  not isdel and sscj="&rscj("levelid")&"")(0)*100,5)&"%</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><a href=lsda.asp?search=gyyy&sscj="&rscj("levelid")&">"&"<font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=0 and sscj="&rscj("levelid")&" and czyy=0")(0)&"</font></a></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=lsda.asp?search=ybyy&sscj="&rscj("levelid")&">"&"<font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=0 and sscj="&rscj("levelid")&"and czyy")(0)&"</font></a></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE fj=1  and not isdel and sscj="&rscj("levelid")&"")(0) &"</div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 and tyzk=1 and sscj="&rscj("levelid")&"")(0) &"</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=lsda.asp?search=3x&sscj="&rscj("levelid")&">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 and tyzk=0 and sscj="&rscj("levelid")&"")(0) &"</a></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 and sscj="&rscj("levelid")&"")(0) &"</div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 and tyzk=1 and sscj="&rscj("levelid")&"")(0) &"</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=lsda.asp?search=2x&sscj="&rscj("levelid")&">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 and tyzk=0 and sscj="&rscj("levelid")&"")(0) &"</a></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=3 and sscj="&rscj("levelid")&"")(0) &"</div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=3 and tyzk=1 and sscj="&rscj("levelid")&"")(0) &"</font></div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=lsda.asp?search=1x&sscj="&rscj("levelid")&">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=3 and tyzk=0 and sscj="&rscj("levelid")&"")(0) &"</a></div></td>"
		dwt.out "    </tr>"
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	
	dwt.out "<tr class=""tdbg"" >" 
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""> ��</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=1 ")(0)&"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda where not isdel ")(0) &"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=1 ")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda where not isdel ")(0)*100,5)&"%</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=0 and czyy=0")(0)&"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and tyzk=0 and czyy")(0)&"</font></div></td>"
	'dwt.out "<td  style=""border-bottom-style: solid;border-width:1px"" >111</td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 ")(0) &"</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 and tyzk=1")(0) &"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 and tyzk=0")(0) &"</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 ")(0) &"</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 and tyzk=1")(0) &"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=2 and tyzk=0")(0) &"</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=3 ")(0) &"</div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=3 and tyzk=1")(0) &"</font></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE not isdel and fj=1 and tyzk=0")(0) &"</div></td>"
	dwt.out "    </tr>"
	dwt.out"</table></div>"
end sub

sub mainbody()
	dwt.out "<SCRIPT language=javascript1.2>" & vbCrLf
	dwt.out "function showsubmenu(sid){" & vbCrLf
	dwt.out "      	 var ss='xxx'+sid;" & vbCrLf
	dwt.out "    whichEl = eval('info' + sid);" & vbCrLf
	dwt.out "    if (whichEl.style.display == 'none'){" & vbCrLf
	dwt.out "        eval(""info"" + sid + "".style.display='block';"");" & vbCrLf
	dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i6.gif' />"";" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "    else{" & vbCrLf
	dwt.out "        eval(""info"" + sid + "".style.display='none';"");" & vbCrLf
	dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i7.gif' />"";" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "}" & vbCrLf
	dwt.out "</SCRIPT>" & vbCrLf

	
	searchs=request("search")
	sqllsda="SELECT * from lsda "
  select case searchs
  case "sscjs"
     url="lsda.asp?sscj="&sscjid&"&search=sscjs"
	 title=" ���� "&sscjh(sscjid)
	 sqllsda=sqllsda&"where sscj="&sscjid
	 if request("action1")="gsk" then 
		 sqllsda=sqllsda&" and fj=1"
		 title=title&" ��˾��"
         url="lsda.asp?sscj="&sscjid&"&search=sscjs&action1=gsk"
	 end if 
	 if request("action1")="ck" then 
		 sqllsda=sqllsda&" and (fj=2 or fj=3)"
		 title=title&" ����"
         url="lsda.asp?sscj="&sscjid&"&search=sscjs&action1=ck"
	 end if 
  case "ssghs"
      url="lsda.asp?ssgh="&ssghid&"&search=ssghs"
	  title=title&" ����װ�� "&gh(ssghid) 
	  sqllsda=sqllsda&"where  ssgh="&ssghid
	 if request("action1")="gsk" then 
		 sqllsda=sqllsda&"  and fj=1"
		 title=title&" ��˾��"
         url="lsda.asp?ssgh="&ssghid&"&search=ssghs&action1=gsk"
	 end if 
	 if request("action1")="ck" then 
		 sqllsda=sqllsda&" AND (fj=2 or fj=3)"
		 title=title&" ����"
         url="lsda.asp?ssgh="&ssghid&"&search=ssghs&action1=ck"
	 end if 
  case "keys"
      url="lsda.asp?keyword="&keys&"&search=keys"
	  title=" ����λ�� '"&keys&" '"
	  sqllsda=sqllsda&"where   wh  like '%" &keys& "%' "
  case "gyyy"
      url="lsda.asp?search=gyyy&sscj="&sscjid
	  title=sscjh(sscjid)&" ����ԭ��δͶ��"
	  sqllsda=sqllsda&"where sscj="&sscjid&" and tyzk=0 and czyy=0"
  case "ybyy"'�༭�ӷ���
      url="lsda.asp?search=gyyy&sscj="&sscjid
	  title=sscjh(sscjid)&" ԭ��δͶ��"
	  sqllsda=sqllsda&"where sscj="&sscjid&" and tyzk=0 and czyy"
  case "3x"
      url="lsda.asp?search=3x&sscj="&sscjid&"&tyzk=0"
	  title=sscjh(sscjid)&" һ��δͶ��"
	  sqllsda=sqllsda&"where sscj="&sscjid&" and tyzk=0 and fj=1"
  case "2x"
      url="lsda.asp?search=2x&sscj="&sscjid&"&tyzk=0"
	  title=sscjh(sscjid)&" ����δͶ��"
	  sqllsda=sqllsda&"where sscj="&sscjid&" and tyzk=0 and fj=2"
  case "1x"
      url="lsda.asp?search=1x&sscj="&sscjid&"&tyzk=0"
	  title=sscjh(sscjid)&" ����δͶ��"
	  sqllsda=sqllsda&"where sscj="&sscjid&" and tyzk=0 and fj=3"
  case "ck"
      url="lsda.asp?search=ck"
	  title="����"
	  sqllsda=sqllsda&"where fj=3 or fj=2"
  case "gsk"
      url="lsda.asp?search=gsk"
	  title="��˾��"
	  sqllsda=sqllsda&"where fj=1"
  case "del"
      url="lsda.asp?search=del"
	  title="��ɾ��"
	  sqllsda=sqllsda&"where isdel ORDER BY deldate deSC"
end select	  	 
	if request("update")<>"" then 
		   'dwt.out sqllsda&" and not isdel ORDER BY update deSC"
		   sqllsda=sqllsda&" where not isdel ORDER BY update deSC"
		   url="lsda.asp?update=update"
	elseif request("search")<>"del" then 
	   sqllsda=sqllsda&" and not isdel ORDER BY SSCJ ASC,ssGH ASC,WH ASC"
'DWT.OUT "DSDFFDS"
	end if 
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>��������"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

  select case searchs
  case "ck"
	call search()
  case "gsk"
	call search()
end select	  	 


'on error resume next
set rslsda=server.createobject("adodb.recordset")
rslsda.open sqllsda,connjg,1,1
if rslsda.eof and rslsda.bof then 
dwt.out "<p align='center'>δ�ҵ���������������</p>" 
else
	dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
	dwt.out "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
	dwt.out "<tr  class=""x-grid-header"">" 
	dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>װ��</div></td>"
	dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>λ��</div></td>"
	dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��;</div></td>"
	dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����ֵL</div></td>"
	dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����ֵH</div></td>"
	if request("search")="del" then 
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��׼����</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��׼�������쵼</div></td>"
    else
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>Ͷ��״��</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>�ּ�</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��ע</div></td>"
    end if 
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>ѡ��</div></td>"
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
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><a href='#' onclick=""showsubmenu("&rslsda("lsdaid")&");"" id=xxx"&rslsda("lsdaid")&"><img src='/img_ext/i7.gif' /></a>"&xh_id&"</div></td>"
                dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_d(rslsda("sscj"))&"</div></td>"
                dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&gh(rslsda("ssgh"))&"</div></td>"
           dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"""
		   if now()-rslsda("update")<7 then   dwt.out "bgcolor=""#FFFF00"""
		        dwt.out ">"&searchH(uCase(rslsda("wh")),keys)&"&nbsp;</td>"
                dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("yt")&"&nbsp;</td>"
                dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("lsl")&"&nbsp;</td>"
                dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("lsh")&"&nbsp;</td>"
         select case rslsda("tyzk")
          case 0
             tyzk="��·"
			 if rslsda("czyy") then
		        tyzk="<font color='#ff0000'>"&tyzk&"</font>"
		      else
		        tyzk="<font color='#0000ff'>"&tyzk&"</font>"
		     end if 	
          case 1 
        	tyzk="<font color='#006600'>Ͷ��</font>"
          
          
        end select	 
		
	if request("search")="del" then 
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("deldate")&"&nbsp;</td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("delr")&"&nbsp;</td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
		if displaypagelevelh(session("groupid"),3,session("pagelevelid")) and displaygrouplevelh(session("groupid"),1,rslsda("sscj")) then
		'dwt.out "<a href=?action=hy&id="&rslsda("lsdaid")&">��ԭ</a>  <a href=?action=deltrue&id="&rslsda("lsdaid")&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">����ɾ��</a></div></td>"
		end if 
		dwt.out "</div>"
    else
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='lsda_czjl.asp?lsdaid="&rslsda("lsdaid")&"&lsdawh="&rslsda("wh")&"'>"&tyzk&"</a>&nbsp;</div></td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&lsdafj(rslsda("fj"))&"</td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rslsda("bz")&"&nbsp;</td>"
		dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"
		'�˾�����ֻ�г��������������������Σ������޸�39
		if  SESSION("USERID")=38 or SESSION("USERID")=39 or SESSION("USERID")=118 or SESSION("USERID")=119 or SESSION("USERID")=110 or SESSION("USERID")=177 then call editdel(rslsda("lsdaid"),rslsda("sscj"),"lsda.asp?action=edit&id=","lsda.asp?action=del&id=")
'130520����Ȩ��ֻ�ܱ༭�Լ����� ��ȷ,���ڱ����������ݵı༭,�༭��Ҫɾ��
 if rslsda("sscj")="1" then call editdel(rslsda("lsdaid"),rslsda("sscj"),"lsda.asp?action=edit&id=","lsda.asp?action=del&id=")
		dwt.out "</div></td>"
    end if 
		
		dwt.out "</tr>"
				
		dwt.out "<tr class='x-grid-row'><td  colspan=7 style='display:none' id='info"&rslsda("lsdaid")&"'>"		
		dwt.out "<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr >" & vbCrLf
		dwt.out "      <td  bgcolor='#BFDFFF'><DIV class='x-grid-hd-text'>һ�μ�����</div></td>"
		dwt.out "      <td  bgcolor='#BFDFFF'><DIV class='x-grid-hd-text'>��λ</div></td>"
		dwt.out "      <td  bgcolor='#BFDFFF'><DIV class='x-grid-hd-text'>��Χ</div></td>"
		dwt.out "      <td  bgcolor='#BFDFFF'><DIV class='x-grid-hd-text'>ִ��װ��</div></td>"
		dwt.out  "    </tr>"
		dwt.out "<tr class='x-grid-row'  >"& vbCrLf
		dwt.out "      <td   bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rslsda("ycjname")&"&nbsp;</td>"
		dwt.out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rslsda("cldw")&"&nbsp;</td>"
		dwt.out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&rslsda("clfw")&"&nbsp;</td>"
		zxzz=rslsda("zxzz")
		'if len(zxzz)>7 then 
		'  zxzz=left(zxzz,6)&"��"
		   
		'	  dwt.out"<script language=javascript src='/js/showPopupText.js'><script>"
		'	  dwt.out "      <td   bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"" onmouseover=""pop('"&rslsda("zxzz")&"','#3366CC');"">"&zxzz&"&nbsp;</td>"
		'else
		  dwt.out "      <td  bgcolor='#BFDFFF' style=""border-bottom-style: solid;border-width:1px"">"&zxzz&"&nbsp;</td>"
		'end if 
        dwt.out  "    </tr>"
		dwt.out "</table>"		
		dwt.out "</tr>"		
				
				
				
				
                 RowCount=RowCount-1
          rslsda.movenext
          loop
        dwt.out "</table>"
       call showpage(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
       rslsda.close
       set rslsda=nothing
        connjg.close
        set connjg=nothing
end sub





dwt.out "</body></html>"

sub search()
	dim sqlcj,rscj
    dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='lsda.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then dwt.out "<a href=""lsda.asp?action=add"">�������</a>"
		dwt.out "&nbsp;&nbsp;<a href='lsda.asp?update=update'>�����������</a>"
		dwt.out "&nbsp;&nbsp;<a href='ls_wh_left.asp'>���޼�¼����</a> "
		'dwt.out "&nbsp;&nbsp;<a href='lsda.asp?search=del'>��ɾ������</a>"
	dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if keys<>"" then 
	 dwt.out "value='"&keys&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='����������λ��'"
	 	dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
	dwt.out "  <input type='hidden' name='search' value='keys'>" & vbCrLf
	
	dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>��������ת����</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='lsda.asp?search=sscjs&sscj="&rscj("levelid")
			if request("search")="gsk"  or request("action1")="gsk" then dwt.out "&action1=gsk"
			if request("search")="ck" or request("action1")="ck"  then dwt.out "&action1=ck"
			dwt.out"'"
			if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"

			dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf
	
	
	dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>��װ����ת����</option>" & vbCrLf
	
	
	
'	sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
'		set rsgh=server.createobject("adodb.recordset")
'		rsgh.open sqlgh,conn,1,1
'		do while not rsgh.eof
'			dwt.out"<option value='lsda.asp?search=ssghs&ssgh="&rsgh("ghid")&"'"
'			if cint(request("ssgh"))=rsgh("ghid") then dwt.out" selected"
'			dwt.out ">"&rsgh("gh_name")&"("&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE ssgh="&rsgh("ghid"))(0)&")</option>"& vbCrLf
'		
'			rsgh.movenext
'		loop
'		rsgh.close
'		set rsgh=nothing
'		dwt.out "     </select>	" & vbCrLf
		
		
		
		
		
	sqlgh="SELECT distinct ssgh,sscj from lsda "
		if request("search")="gsk"  or request("action1")="gsk" then sql=sql&" where fj=1"
		if request("search")="ck" or request("action1")="ck"  then sql=sql&" where fj=2 or fj=3"
	sqlgh=sqlgh&" order by sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connjg,1,1
    do while not rsgh.eof
		ghid=cint(rsgh("ssgh"))


		sql="SELECT count(lsdaid) FROM lsda WHERE ssgh="&ghid
		if request("search")="gsk" or request("action1")="gsk" then sql=sql&" and fj=1"
		if request("search")="ck" or request("action1")="ck" then sql=sql&" and ( fj=2 or fj=3 )"
		sb_numb=Connjg.Execute(sql)(0)
        
		if sb_numb<>0 then 
			i=i+1
			dwt.out"<option value='lsda.asp?search=ssghs&ssgh="&ghid
			if request("search")="gsk" or request("action1")="gsk"  then dwt.out "&action1=gsk"
			if request("search")="ck"  or request("action1")="ck" then dwt.out "&action1=ck"

			dwt.out "'"
			if cint(request("ssgh"))=ghid then dwt.out" selected"
			
			sql="SELECT gh_name FROM ghname WHERE ghid="&ghid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    gh_name="δ֪��"
			else
			    gh_name=rs("gh_name")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&i&"&nbsp;&nbsp;"&gh_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf
		
		
		
		
		
		
		
dwt.out "������ǰ+����ʾ��ϸ��Ϣ</form></div></div>" & vbCrLf


'	'dwt.out"<select name='lsda_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
'   dwt.out"<select name='lsda_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
'    dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'    dwt.out"</select>"  	 & vbCrLf
'    'dwt.out "<select name='lsda_gh' size='1'  onChange=""alert(document.all.lsda_sscj.options[document.all.lsda_sscj.selectedIndex].value);alert(this.value);"">" & vbCrLf
'	dwt.out "<select name='lsda_gh' size='1' onChange=""location.href='lsda_search.asp?action=sscjs&sscj=' + document.all.lsda_sscj.options[document.all.lsda_sscj.selectedIndex].value + '&ssgh=' + this.value;"">" & vbCrLf
'    dwt.out "<option  selected>ѡ��װ�÷���</option>" & vbCrLf
'    dwt.out "</select></td></tr>  "  & vbCrLf
'    dwt.out "<script><!--" & vbCrLf
'    dwt.out "var groups=document.all.lsda_sscj.options.length" & vbCrLf
'    dwt.out "var group=new Array(groups)" & vbCrLf
'    dwt.out "for (i=0; i<groups; i++)" & vbCrLf
'    dwt.out "group[i]=new Array()" & vbCrLf
'    dwt.out "group[0][0]=new Option(""ѡ��װ�÷���"","" "");" & vbCrLf
'	
'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    if rscj.eof then 
'	  dwt.out "δ�ҵ�����"
'    else
'	do while not rscj.eof
'     lsdaii=1		
'		sqlgh="SELECT * from ghname where sscj="&rscj("levelid")
'        set rsgh=server.createobject("adodb.recordset")
'        rsgh.open sqlgh,conn,1,1
'        if rsgh.eof and rsgh.bof then
'		   dwt.out "group["&rscj("levelid")&"][0]=new Option(""δ���װ��"",""0"");" & vbCrLf
'		else
'		   dwt.out"group["&rsgh("sscj")&"][0]=new Option(""ѡ��װ�÷���"",""0"");" & vbCrLf
'		do while not rsgh.eof
'		   dwt.out"group["&rsgh("sscj")&"]["&lsdaii&"]=new Option("""&rsgh("gh_name")&""","""&rsgh("ghid")&""");" & vbCrLf
'		  lsdaii=lsdaii+1
'		   rsgh.movenext
'	    loop
'	    end if 
'		rsgh.close
'	    set rsgh=nothing
'
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'
'  end if 
'
'
'    dwt.out "var temp=document.all.lsda_gh" & vbCrLf
'    dwt.out "function redirect(x){" & vbCrLf
'    dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
'    dwt.out "temp.options[m]=null" & vbCrLf
'    dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
'    dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
'    dwt.out "}" & vbCrLf
'    dwt.out "temp.options[0].selected=true" & vbCrLf
''    dwt.out "location.href=""lsda_search.asp?action=sscjs&sscj=""+x + ""&ssgh="" + group[x][0].value"
'	dwt.out "}//-->" & vbCrLf ��JS������־
end sub


function lsdafj(fjnumb)
	if isnull(fjnumb) or fjnumb=0 then 
	  lsdafj="δ�ּ�"
	else
		'for fj_i=1 to fjnumb
		'fj=fj&"*"
		'next
	  if fjnumb=1 then lsdafj="����"
	  if fjnumb=2 then lsdafj="���"
	  if fjnumb=3 then lsdafj="��"
	end if 
end function 



Call Closeconn
%>