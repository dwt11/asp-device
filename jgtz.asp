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
dim sqljgtz,rsjgtz,title,record,pgsz,total,page,start,rowcount,url,ii,zxzz,jx_numb
dim rsadd,sqladd,jgtzid,rsedit,sqledit,scontent,rsdel,sqldel,tyzk,id,wcyear,sql
url="jgtz.asp"
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

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>��Ϣ����ϵͳ����̨�˹���ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form1.jgtz_sscj.value==''){" & vbCrLf
dwt.out "      alert('��ѡ���������䣡');" & vbCrLf
dwt.out "   document.form1.jgtz_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

sub add()
dim rscj,sqlcj

   dwt.out"<br><br><form method='post' action='jgtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>��Ӽ���̨��</strong></div></td>    </tr>"
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='jgtz_sscj' size='1'>"
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
     dwt.out"<input name='jgtz_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
      dwt.out"<input name='jgtz_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf

 end if 

	 
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>��Ŀ���ƣ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='jgtz_name' type='text'  size=""50""></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ˣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_tcr'></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_tcdate' type='text' value="&now()&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_tcdate, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ĿͶ�ʣ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_xmtz'></td></tr>  "   
   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��Լ���ʽ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_jyjjz'></td></tr>  "   

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ԭ��</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"   
	dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=jgtz_jgyy&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
    dwt.out"</iframe><textarea name='jgtz_jgyy' style='display:none'></textarea>"
    dwt.out"</td></tr>"   

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ķ�����</strong></td>"      
	dwt.out"<td width='88%' class='tdbg'>"   
	dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=jgtz_jgfa&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
    dwt.out"</iframe><textarea name='jgtz_jgfa' style='display:none'></textarea>"
    dwt.out"</td></tr>"   

	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_pf_date' type='text'>"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_pf_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_pf_qk'></td></tr>  "   



	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_ssdate' type='text'>"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_ssdate, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩ�����ˣ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_ssname'></td></tr>  "   
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩ�����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><textarea name=""jgtz_ssqk"" cols=""50"" rows=""15""></textarea></td></tr>  "   
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����Ч����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><textarea name=""jgtz_jgxg"" cols=""50"" rows=""15""></textarea></td></tr>  "   



	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_wc_date' type='text'>"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_wc_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��������</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_wc_qk'></td></tr>  "   
	
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_jd_date' type='text'>"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_jd_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�����ȼ���</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_jd_dj'></td></tr>  "   
	



	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
	  '����
     on error resume next
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from jgtz" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("jgtz_sscj"))
      rsadd("name")=request("jgtz_name")
      	  rsadd("tcr")=request("jgtz_tcr")
      rsadd("tcdate")=Trim(request("jgtz_tcdate"))
      rsadd("xmtz")=request("jgtz_xmtz")
      rsadd("jyjjz")=request("jgtz_jyjjz")
      rsadd("jgyy")=request("jgtz_jgyy")
      rsadd("jgfa")=request("jgtz_jgfa")
      rsadd("ssdate")=request("jgtz_ssdate")
      rsadd("ssname")=request("jgtz_ssname")
      rsadd("ssqk")=request("jgtz_ssqk")
      rsadd("jgxg")=request("jgtz_jgxg")
	  
	        rsadd("pf_qk")=request("jgtz_pf_qk")
      rsadd("pf_date")=request("jgtz_pf_date")
      rsadd("wc_qk")=request("jgtz_wc_qk")
      rsadd("wc_date")=request("jgtz_wc_date")
	  rsadd("wc_year")=cint(year(request("jgtz_wc_date")))
      rsadd("jd_date")=request("jgtz_jd_date")
      rsadd("jd_dj")=request("jgtz_jd_dj")


      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='jgtz.asp';;</Script>"
end sub
sub del()
  jgtzid=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from jgtz where id="&jgtzid
  rsdel.open sqldel,connjg,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub saveedit()    
	 
	 on error resume next '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jgtz where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
      rsedit("sscj")=Trim(Request("jgtz_sscj"))
      rsedit("name")=request("jgtz_name")
	  rsedit("tcr")=request("jgtz_tcr")
      rsedit("tcdate")=Trim(request("jgtz_tcdate"))
      rsedit("xmtz")=request("jgtz_xmtz")
      rsedit("jyjjz")=request("jgtz_jyjjz")
      rsedit("jgyy")=request("jgtz_jgyy")
      rsedit("jgfa")=request("jgtz_jgfa")
      rsedit("ssdate")=request("jgtz_ssdate")
      rsedit("ssname")=request("jgtz_ssname")
      rsedit("ssqk")=request("jgtz_ssqk")
      rsedit("jgxg")=request("jgtz_jgxg")
      
      rsedit("pf_qk")=request("jgtz_pf_qk")
      rsedit("pf_date")=request("jgtz_pf_date")
      rsedit("wc_qk")=request("jgtz_wc_qk")
      rsedit("wc_date")=request("jgtz_wc_date")
	  rsedit("wc_year")=cint(year(request("jgtz_wc_date")))
      rsedit("jd_date")=request("jgtz_jd_date")
      rsedit("jd_dj")=request("jgtz_jd_dj")

	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub



sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from jgtz where id="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><form method='post' action='jgtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�༭����̨��</strong></div></td>    </tr>"
     
     dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     dwt.out"<td width='88%' class='tdbg'><input name='jgtz_sscj'  disabled='disabled'  type='text' value='"&sscjh_d(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out"<input name='jgtz_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>��Ŀ���ƣ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='jgtz_name' type='text'  size=""50"" value='"&rsedit("name")&"'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ˣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_tcr'  value='"&rsedit("tcr")&"'></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_tcdate' type='text' value="&rsedit("tcdate")&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_tcdate, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ĿͶ�ʣ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_xmtz' value='"&rsedit("xmtz")&"'></td></tr>  "   
   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��Լ���ʽ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_jyjjz' value='"&rsedit("jyjjz")&"'></td></tr>  "   

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ԭ��</strong></td>"      
    scontent=rsedit("jgyy")
	dwt.out"<td width='88%' class='tdbg'>"   
	dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=jgtz_jgyy&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
    dwt.out"</iframe><textarea name='jgtz_jgyy' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    dwt.out"</td></tr>"   

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ķ�����</strong></td>"      
    scontent=rsedit("jgfa")
	dwt.out"<td width='88%' class='tdbg'>"   
	dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=jgtz_jgfa&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
    dwt.out"</iframe><textarea name='jgtz_jgfa' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    dwt.out"</td></tr>"   

	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_pf_date' type='text' value="&rsedit("pf_date")&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_pf_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_pf_qk' value='"&rsedit("pf_qk")&"'></td></tr>  "   



	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_ssdate' type='text' value="&rsedit("ssdate")&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_ssdate, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩ�����ˣ�</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_ssname' value='"&rsedit("ssname")&"'></td></tr>  "   
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵʩ�����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><textarea name=""jgtz_ssqk"" cols=""50"" rows=""15"">"&rsedit("ssqk")&"</textarea></td></tr>  "   
	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����Ч����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><textarea name=""jgtz_jgxg"" cols=""50"" rows=""15"">"&rsedit("jgxg")&"</textarea></td></tr>  "   



	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_wc_date' type='text' value="&rsedit("wc_date")&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_wc_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��������</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_wc_qk' value='"&rsedit("wc_qk")&"'></td></tr>  "   
	
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='jgtz_jd_date' type='text' value="&rsedit("jd_date")&" >"
   dwt.out"<a href='#' onClick=""popUpCalendar(this, jgtz_jd_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�����ȼ���</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='jgtz_jd_dj' value='"&rsedit("jd_dj")&"'></td></tr>  "   
	


	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

end sub


sub main()
dwt.out "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out " <tr class='topbg'>"& vbCrLf
dwt.out "   <td height='22' colspan='2' align='center'><strong>����̨�˹���ҳ</strong></td>"& vbCrLf
dwt.out "  </tr>  "& vbCrLf
dwt.out "<tr class='tdbg'>"& vbCrLf
dwt.out "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
dwt.out "    <td height='30'><a href=""jgtz.asp"">����̨����ҳ</a>&nbsp;|&nbsp;<a href=""jgtz.asp?action=add"">��Ӽ�����Ŀ</a></td>"& vbCrLf
dwt.out "  </tr>"& vbCrLf
dwt.out "</table>"& vbCrLf
call search()
sqljgtz="SELECT * from jgtz ORDER BY tcdate DESC"
set rsjgtz=server.createobject("adodb.recordset")
rsjgtz.open sqljgtz,connjg,1,1
if rsjgtz.eof and rsjgtz.bof then 
dwt.out "<p align='center'>δ��Ӽ�����Ŀ</p>" 
else

dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out "<tr class=""title"">" 
dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center""><strong>���</strong></div></td>"
dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center""><strong>ʵʩ����</strong></div></td>"
dwt.out "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��Ŀ����</strong></div></td>"
dwt.out "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�����</strong></div></td>"
dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���ʱ��</strong></div></td>"
dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ʵʩʱ��</strong></div></td>"
dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���ʱ��</strong></div></td>"
dwt.out "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"
dwt.out "    </tr>"
           record=rsjgtz.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsjgtz.PageSize = Cint(PgSz) 
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
           rsjgtz.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsjgtz.PageSize
           do while not rsjgtz.eof and rowcount>0
                 dwt.out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center"">"&rsjgtz("id")&"</div></td>"
                dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center"">"&sscjh(rsjgtz("sscj"))&"</div></td>"
                dwt.out "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px""><a href=jgtz_view.asp?id="&rsjgtz("id")&">"&rsjgtz("name")&"</a>&nbsp;</div></td>"
                dwt.out "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("tcr")&"&nbsp;</div></td>"
                dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("tcdate")&"&nbsp;</div></td>"
                dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("ssdate")&"&nbsp;</div></td>"
				dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("wc_date")&"&nbsp;</div></td>"
                dwt.out "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				call editdel(rsjgtz("id"),rsjgtz("sscj"),"jgtz.asp?action=edit&id=","jgtz.asp?action=del&id=")
				
                dwt.out "</div></td></tr>"
                 RowCount=RowCount-1
          rsjgtz.movenext
          loop
        dwt.out "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rsjgtz.close
       set rsjgtz=nothing
        connjg.close
        set connjg=nothing
end sub





dwt.out "</body></html>"


sub search()
dim rscj,sqlcj
dwt.out "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='jgtz_search.asp'>" & vbCrLf
dwt.out "  <tr class='tdbg'>   <td>" & vbCrLf
dwt.out "  <strong>��Ŀ������</strong>" & vbCrLf
dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"">" & vbCrLf
dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
dwt.out "  <input type='hidden' name='Action' value='keys'>" & vbCrLf
dwt.out "</td></form><td><font color='0066CC'> �鿴���������������ݣ�</font>"
dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
dwt.out "	       <option value=''>��������ת����</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='jgtz_search.asp?action=sscjs&sscj="&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	dwt.out "     </select>	" & vbCrLf
	
dwt.out "<select name='Jump2Class' id='Jump2Class' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>����������ת����</option>" & vbCrLf
	sqljgtz="SELECT distinct wc_year from jgtz"
	set rsjgtz=server.createobject("adodb.recordset")
    rsjgtz.open sqljgtz,connjg,1,1
    do while not rsjgtz.eof
		
		sql="SELECT count(id) FROM jgtz WHERE wc_year like'%"&rsjgtz("wc_year")&"%'"
		jx_numb=Connjg.Execute(sql)(0)
        
		if jx_numb<>0 then 
			Dwt.out"<option  value='jgtz_search.asp?action=wcyears&wcyear="&rsjgtz("wc_year")&"'"
			if request("wcyear")=rsjgtz("wc_year") then Dwt.out" selected"
			Dwt.out ">"&rsjgtz("wc_year")&"("&jx_numb&")</option>"& vbCrLf '
	    end if 

		rsjgtz.movenext
	loop
	rsjgtz.close
	set rsjgtz=nothing
	Dwt.out "     </select>	" & vbCrLf

	
dwt.out "	</td>  </tr></table>" & vbCrLf
end sub





Call Closeconn
%>