<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->


<%
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ���ܽ�������ʾ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checkadd(){" & vbCrLf
response.write " if(document.form1.yzj_sscj.value==''){" & vbCrLf
response.write "      alert('��ѡ���������䣡');" & vbCrLf
response.write "   document.form1.yzj_sscj.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf

response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf

response.write "   <td height='22' colspan='2' align='center'><strong>��֧�����ܽ�</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""d_yjhzj.asp"">�¼ƻ��ܽ������ҳ</a>&nbsp;|&nbsp;<a href=""d_yjh_view.asp?action=addyjh"">����¼ƻ�</a>&nbsp;|&nbsp;<a href=""d_yzj_view.asp?action=addyzj"">������ܽ�</a></td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf
if request("action")="yzj" then call yzj()
if request("action")="addyzj" then call addyzj()
if request("action")="saveaddyzj" then call saveaddyzj()
if request("action")="del" then call del()
if request("action")="edit" then call edit()
if request("action")="saveedit" then saveedit()
'���з������˿����޸��Լ�������
sub addyzj()
dim ii
dim rscj,sqlcj,rsbz,sqlbz,sql,rs
   response.write"<form method='post' action='d_yzj_view.asp' name='formyzj'  onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>������ܽ�</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>����֧���� </strong></td>"& vbCrLf      
    response.write"<td width='88%' class='tdbg'>"& vbCrLf
  'if session("level")=0 then 
	response.write"<select name='yzj_sscj' size='1'>"
    response.write"<option >��ѡ������֧��</option>"
		response.write"<option  value=1>ά��һ��֧��</option>"
		response.write"<option  value=2>ά�޶���֧��</option>"
		response.write"<option  value=3>���ص�֧��</option>"
		response.write"<option  value=4>ά������֧��</option>"
		response.write"<option  value=5>ά���ĵ�֧��</option>"



'    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
    response.write"</select></td></tr>  "  	 
'  else 	 
'    response.write"<input name='yzj_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
'    response.write"<input name='yzj_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
' end if 
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ܽ����ڣ�</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='yzj_date' type='text' value="&year(now())&"-"&month(now())&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, yzj_date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>���ܽ����ݣ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'>"
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=yzj_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe><textarea name='yzj_body' style='display:none'></textarea>"
	  response.write"</td></tr>   "
	 
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' value='saveaddyzj'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveaddyzj()    
	  'response.write"DDDDDDDDDDDDDDDDDDDDD"
	  dim year1,month1,day1'����\
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from d_yzj" 
      rsadd.open sqladd,conna,1,3
      rsadd.addnew
      rsadd("sscj")=Request("yzj_sscj")
      year1=year(Trim(Request("yzj_date")))
	  month1=month(Trim(Request("yzj_date")))
	  if len(month1)<>2 then month1="0"&month1
      rsadd("month")=month1
	  rsadd("year")=year1
	  dim yzj_body
      yzj_body=request("yzj_body")
	  if yzj_body="" then   yzj_body="δ�������"
	  rsadd("body")=yzj_body
      rsadd("userid")=session("userid")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='d_yjhzj.asp';</Script>"
end sub



sub saveedit()    
	  dim year1,month1,day1'����\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from d_yzj where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conna,1,3
      rsedit("sscj")=Trim(Request("yzj_sscj"))
      year1=year(Trim(Request("yzj_date")))
	  month1=month(Trim(Request("yzj_date")))
	  if len(month1)<>2 then month1="0"&month1
      rsedit("month")=month1
	  rsedit("year")=year1
      rsedit("body")=Trim(request("yzj_body"))

      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub edit()
    dim scontent
   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from d_yzj where id="&id
   rsedit.open sqledit,conna,1,1

   response.write"<form method='post' action='d_yzj_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�༭���ܽ�</strong></div></td>    </tr>"
	
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����֧���� </strong></td>"   & vbCrLf   
     
	  		dim sszb
		if rsedit("sscj")=1 then sszb="ά��һ��֧��"
		if rsedit("sscj")=2 then sszb="ά�޶���֧��"
		if rsedit("sscj")=3 then sszb="���ص�֧��"
		if rsedit("sscj")=4 then sszb="ά������֧��"
		if rsedit("sscj")=5 then sszb="ά���ĵ�֧��"
	


	 response.write"<td width='88%' class='tdbg'><input name='yzj_sscj'  disabled='disabled'  type='text' value='"&sszb&"'></td></tr>"& vbCrLf
     response.write"<input name='yzj_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf
	    
		response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ܽ����ڣ�</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='yzj_date' type='text' value="&rsedit("year")&"-"&rsedit("month")&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, yzj_date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ݣ� </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
	scontent=rsedit("body")
	 response.write"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=yzj_body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      response.write"</iframe><textarea name='yzj_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    response.write"</td></tr>  "   

	
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub

sub del()
 dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from d_yzj where id="&request("id")
  rsdel.open sqldel,conna,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub yzj()
   dim rsyzj,sqlyzj,rs,sql
   '��ʾ���伶�����ܽ��ܽ�
      sqlyzj="SELECT * from d_yzj where sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rsyzj=server.createobject("adodb.recordset")
      rsyzj.open sqlyzj,conna,1,1
    	  		dim sszb
		if request("sscj")=1 then sszb="ά��һ��֧��"
		if request("sscj")=2 then sszb="ά�޶���֧��"
		if request("sscj")=3 then sszb="���ص�֧��"
        if request("sscj")=4 then sszb="ά������֧��"
		if request("sscj")=5 then sszb="ά���ĵ�֧��"

		 response.write "<br><table width=""90%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			 
             response.write " <tr class=""title""><td height=40><div align=center><strong>"&sszb&request("year")&"��"&request("month")&"�·ݹ����ܽ�</strong></div>"
             response.write "</td></tr>"
             response.write "<tr class=""tdbg"">"
			 response.write "<td><table width=""90%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"
			 response.write rsyzj("body")
			 response.write "</td></tr></table></td>"
             response.write "</tr></table><br>"		
  rsyzj.close
  set rsyzj=nothing
end sub


response.write "</body></html>"
Call CloseConn
%>