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
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ��ѵ����������ʾ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checkadd(){" & vbCrLf
response.write " if(document.form1.pxzj_sscj.value==''){" & vbCrLf
response.write "      alert('��ѡ���������䣡');" & vbCrLf
response.write "   document.form1.pxzj_sscj.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write " if(document.form1.pxzj_numb.value==''){" & vbCrLf
response.write "      alert('����дӦ���˴Σ�');" & vbCrLf
response.write "   document.form1.pxzj_numb.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write " if(document.form1.pxzj_sdnumb.value==''){" & vbCrLf
response.write "      alert('����дʵ���˴Σ�');" & vbCrLf
response.write "   document.form1.pxzj_sdnumb.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write " if(document.form1.pxzj_hgnumb.value==''){" & vbCrLf
response.write "      alert('����д�ϸ��˴Σ�');" & vbCrLf
response.write "   document.form1.pxzj_hgnumb.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "        //�ж�����(�����ڸ�ʽΪ:2006-04-03 17:55:00)������" & vbCrLf
response.write "var dval1=document.form1.pxzj_date.value;" & vbCrLf
response.write "var r=/^(\d{0,4})-(0{0,1}[1-9]|1[0-2])-(0{0,1}[1-9]|[1-2]\d|3[0-1])$/;" & vbCrLf
response.write "if(!r.test(dval1)){" & vbCrLf
response.write "    alert('�������ڴ���');" & vbCrLf
response.write "    document.form1.pxzj_date.focus();" & vbCrLf
response.write "    return false;}" & vbCrLf
response.write "else{" & vbCrLf
response.write "    var r1=/^0{0,4}$/;" & vbCrLf
response.write "    if(r1.test(RegExp.$1)){ " & vbCrLf
response.write "           alert('��ݲ���Ϊ0');" & vbCrLf
response.write "             document.form1.pxzj_date.focus();" & vbCrLf
response.write "              return false;}" & vbCrLf
response.write "              var r2=/1[02]|0{0,1}[13578]/;" & vbCrLf' //С��
response.write "              if(!r2.test(RegExp.$2)){" & vbCrLf
response.write "                     if(parseInt(RegExp.$2)==2){" & vbCrLf
response.write "                            if(parseInt(RegExp.$1)%4==0){" & vbCrLf
response.write "                                 if(parseInt(RegExp.$3)>29){" & vbCrLf
response.write "                                       alert('����2��ֻ��29��');" & vbCrLf
response.write "                                       document.form1.pxzj_date.focus();" & vbCrLf
response.write "                                       return false;}}" & vbCrLf
response.write "                             else{" & vbCrLf
response.write "                                 if(parseInt(RegExp.$3)>28){" & vbCrLf
response.write "                                       alert('2��ֻ��28��');" & vbCrLf
response.write "                                       document.form1.pxzj_date.focus();" & vbCrLf
response.write "                                       return false;}}}" & vbCrLf
response.write "                             else{" & vbCrLf
response.write "                                if(parseInt(RegExp.$3)>30){" & vbCrLf
response.write "                                        alert('С��ֻ��30��');" & vbCrLf
response.write "                                        document.form1.pxzj_date.focus();" & vbCrLf
response.write "                                        return false;}}}}" & vbCrLf


response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if request("action")="pxzj" then call pxzj()
if request("action")="addpxzj" then call addpxzj()
if request("action")="saveaddpxzj" then call saveaddpxzj()
if request("action")="del" then call del()
if request("action")="edit" then call edit()
if request("action")="saveedit" then saveedit()
sub addpxzj()
dim ii
dim rscj,sqlcj,rsbz,sqlbz,sql,rs
   response.write"<br><br><br><form method='get' action='pxzj_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�����ѵ�ܽ�</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"& vbCrLf      
    response.write"<td width='80%' class='tdbg'>"& vbCrLf
  if session("level")=0 then 
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	response.write"<select name='pxzj_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    response.write"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select>"  	 & vbCrLf
    response.write "<select name='pxzj_ssbz' size='1' >" & vbCrLf
    response.write "<option  selected>ѡ��������</option>" & vbCrLf
    response.write "</select></td></tr>  "  & vbCrLf
    response.write "<script><!--" & vbCrLf
    response.write "var groups=document.form1.pxzj_sscj.options.length" & vbCrLf
    response.write "var group=new Array(groups)" & vbCrLf
    response.write "for (i=0; i<groups; i++)" & vbCrLf
    response.write "group[i]=new Array()" & vbCrLf
    response.write "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   response.write "group["&rscj("levelid")&"][0]=new Option(""����"",""0"");" & vbCrLf
		else
		   response.write"group["&rscj("levelid")&"][0]=new Option(""����"",""0"");" & vbCrLf
		do while not rsbz.eof
		   response.write"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

		rscj.movenext
	loop
	rscj.close
	set rscj=nothing




    response.write "var temp=document.form1.pxzj_ssbz" & vbCrLf
    response.write "function redirect(x){" & vbCrLf
    response.write "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    response.write "temp.options[m]=null" & vbCrLf
    response.write "for (i=0;i<group[x].length;i++){" & vbCrLf
    response.write "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    response.write "}" & vbCrLf
    response.write "temp.options[0].selected=true" & vbCrLf
    response.write "}//--></script>" & vbCrLf



  else 	 
   response.write"<input name='pxzj_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   response.write"<input name='pxzj_sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   response.write"<select name='pxzj_ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  response.write"<option value='0'>����</option>"
   else   
	  response.write"<option value='0'>����</option>"
      do while not rs.eof
	     response.write"<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	  end if 
	 response.Write"</select>" 
  rs.close
  set rs=nothing
 end if 
    response.write"</td></tr>  "  	 

	 	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��ѵ�ܽ����ڣ�</strong></td> "
   response.write"<td width='80%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='pxzj_date' type='text' value="&now()&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, pxzj_date, 'yyyy-mm-dd'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>��ѵ����ժҪ��</strong></td>"
	 'response.write"<td width='80%' class='tdbg'><input name='pxzj_body' type='text'  size=""50""></td>    </tr>   "
	 response.write"<td width='80%' class='tdbg'>"
	 sql="SELECT * from pxjh where sscj="&session("levelclass")&" and year="&year(now())&" and month="&month(now())
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conne,1,1
   response.write"<select name='pxzj_body' size='1'>"
   
   if rs.eof and rs.bof then 
   	  response.write"����δ��Ӽƻ�"
   else   
	  'response.write"<option value='0'>����</option>"
      do while not rs.eof
	     response.write"<option value='"&rs("body")&"'>"&rs("body")&"</option>"
	  rs.movenext
      loop
	  end if 
	 response.Write"</select>" 
  rs.close
  set rs=nothing

	 response.write"</td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ����</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_dx'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>Ӧ���˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_numb'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵ���˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_sdnumb'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ϸ��˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_hgnumb'></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ��ʽ��</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_xs'></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ��ʱ��</strong></td> "
    response.write"<td width='80%' class='tdbg'>"
	response.write"<select name='pxzj_ks' size='1'>"
	response.write"<option value='1'>1h</option>"
    response.write"<option value='2'>2h</option>"
    response.write"<option value='3'>3h</option>"
    response.write"<option value='4'>4h</option>"
    response.write"<option value='5'>5h</option>"
    response.write"</select></td></tr>  "  	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ڿ��ˣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_skrname'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ע��</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_bz'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ˣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_tbrname'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��λ���ܣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_zgname'></td></tr> "
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveaddpxzj'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveaddpxzj()    
	  dim year1,month1,day1'����\
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from pxzj" 
      rsadd.open sqladd,conne,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("pxzj_sscj"))
      rsadd("ssbz")=Trim(Request("pxzj_ssbz"))
      year1=year(Trim(Request("pxzj_date")))
	  month1=month(Trim(Request("pxzj_date")))
	  day1=day(Trim(Request("pxzj_date")))
	  if len(month1)<>2 then month1="0"&month1
	  rsadd("day")=day1
      rsadd("month")=month1
	  rsadd("year")=year1
	  rsadd("tbrname")=request("pxzj_tbrname")
	  rsadd("zgname")=request("pxzj_zgname")
      rsadd("body")=request("pxzj_body")
	  rsadd("dx")=request("pxzj_dx")
	  rsadd("numb")=request("pxzj_numb")
	  rsadd("sdnumb")=request("pxzj_sdnumb")
	  rsadd("hgnumb")=request("pxzj_hgnumb")
      rsadd("xs")=request("pxzj_xs")
      rsadd("ks")=request("pxzj_ks")
      rsadd("skrname")=request("pxzj_skrname")
      rsadd("tbdate")=year(now())&"-"&month(now())&"-"&day(now())
      rsadd("bz")=request("pxzj_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>history.go(-2);</Script>"
end sub
sub saveedit()    
	  dim year1,month1,day1'����\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from pxzj where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conne,1,3
      rsedit("sscj")=Trim(Request("pxzj_sscj"))
      rsedit("ssbz")=Trim(Request("pxzj_ssbz"))
      year1=year(Trim(Request("pxzj_date")))
	  month1=month(Trim(Request("pxzj_date")))
	  day1=day(Trim(Request("pxzj_date")))
	  if len(month1)<>2 then month1="0"&month1
	  rsedit("day")=day1
      rsedit("month")=month1
	  rsedit("year")=year1
	  rsedit("tbrname")=request("pxzj_tbrname")
	  rsedit("zgname")=request("pxzj_zgname")
      rsedit("body")=request("pxzj_body")
	  rsedit("dx")=request("pxzj_dx")
	  rsedit("numb")=request("pxzj_numb")
	  rsedit("sdnumb")=request("pxzj_sdnumb")
	  rsedit("hgnumb")=request("pxzj_hgnumb")
      rsedit("xs")=request("pxzj_xs")
      rsedit("ks")=request("pxzj_ks")
      rsedit("skrname")=request("pxzj_skrname")
      rsedit("tbdate")=year(now())&"-"&month(now())&"-"&day(now())
      rsedit("bz")=request("pxzj_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub edit()

   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from pxzj where id="&id
   rsedit.open sqledit,conne,1,1

   response.write"<br><br><br><form method='get' action='pxzj_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='20' colspan='2'>"
   response.write"<div align='center'><strong>�༭��ѵ�ܽ�</strong></div></td>    </tr>"
	
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"& vbCrLf      
    response.write"<td width='80%' class='tdbg'>"& vbCrLf
    response.write"<input name=""pxzj_sscj"" value="&sscjh(rsedit("sscj"))&" type='text' disabled='disabled' >"& vbCrLf
     response.write"<input name='pxzj_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������飺 </strong></td>"& vbCrLf      
    response.write"<td width='80%' class='tdbg'>"& vbCrLf
	if rsedit("ssbz")=0 then
  	   ssbz="����"
	else
	   ssbz=ssbzh(rsedit("ssbz"))
	end if    
    response.write"<input name=""pxzj_ssbz"" value="&ssbz&" type='text' disabled='disabled' >"& vbCrLf
     response.write"<input name='pxzj_ssbz' type='hidden' value="&rsedit("ssbz")&"></td></tr>"& vbCrLf

   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ�ƻ����ڣ�</strong></td> "
   response.write"<td width='80%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='pxzj_date' type='text' value="&rsedit("year")&"-"&rsedit("month")&"-"&rsedit("day")&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this, pxzj_date, 'yyyy-mm-dd'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>��ѵ����ժҪ��</strong></td>"
	 response.write"<td width='80%' class='tdbg'><input name='pxzj_body' type='text'  size=""50"" value='"&rsedit("body")&"'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ����</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_dx' value='"&rsedit("dx")&"'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>Ӧ���˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_numb' value='"&rsedit("numb")&"'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ʵ���˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_sdnumb' value='"&rsedit("sdnumb")&"'></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ϸ��˴Σ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_hgnumb' value='"&rsedit("hgnumb")&"'></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ��ʽ��</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_xs' value="&rsedit("xs")&"></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ѵ��ʱ��</strong></td> "
    response.write"<td width='80%' class='tdbg'>"
	response.write"<select name='pxzj_ks' size='1'>"
	
	response.write"<option value='1'"
	if rsedit("ks")=1 then response.write"selected"
	response.write">1h</option>"
	
    response.write"<option value='2'"
	if rsedit("ks")=2 then response.write"selected"
	response.write">2h</option>"
	
    response.write"<option value='3'"
	if rsedit("ks")=3 then response.write"selected"
	response.write">3h</option>"
	
    response.write"<option value='4'"
	if rsedit("ks")=4 then response.write"selected"
	response.write">4h</option>"
	
    response.write"<option value='5"
	if rsedit("ks")=5 then response.write"selected"
	response.write"'>5h</option>"
    
	response.write"</select></td></tr>  "  	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ڿ��ˣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_skrname' value="&rsedit("skrname")&"></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ע��</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_bz' value="&rsedit("bz")&"></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ˣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_tbrname' value="&rsedit("tbrname")&"></td></tr> "
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��λ���ܣ�</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='pxzj_zgname' value="&rsedit("zgname")&"></td></tr> "
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub

sub del()
 dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from pxzj where id="&request("id")
  rsdel.open sqldel,conne,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub pxzj()

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf

response.write "   <td height='22' colspan='2' align='center'><div align=center><strong>"&sscjh(request("sscj"))&request("year")&"��"&request("month")&"�·���ѵ�ܽ�</strong></div></strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf

   dim rspxzj,sqlpxzj,rs,sql
   '��ʾ���伶����ѵ�ܽ�
      sqlpxzj="SELECT * from pxzj where ssbz=0 and sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rspxzj=server.createobject("adodb.recordset")
      rspxzj.open sqlpxzj,conne,1,1
      if rspxzj.eof and rspxzj.bof then 
             response.write "<p align='center'>δ��ӳ�����ѵ�ܽ�</p>" 
          else
             response.write "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
             response.write " <tr class=""title""><td colspan=13 >&nbsp;&nbsp;&nbsp;  ��λ��"&sscjh(request("sscj"))&"&nbsp;"&ssbzh(rspxzj("ssbz"))
             response.write "</td></tr>"
             response.write "<tr class=""title""><td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ʱ��</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ����ժҪ</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ����</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>Ӧ���˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ʵ���˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ϸ��˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ��</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ϸ���</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ��ʽ</div></td>"
             response.write " <td   style=""border-bottom-style: solid;border-width:1px""><div align=center>�ۼƿ�ʱ</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ڿ���</div></td>"
             response.write " <td   style=""border-bottom-style: solid;border-width:1px""><div align=center>��ע</div></td>"
			 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ѡ��</div></td></tr>"
              do while not rspxzj.eof
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
				 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("month")&"."&rspxzj("day")&"</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px"">"&rspxzj("body")&"&nbsp;</td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("dx")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("numb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("sdnumb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("hgnumb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("sdnumb")/rspxzj("numb")*100&"%&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("hgnumb")/rspxzj("numb")*100&"%&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("xs")&"&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("ks")&"h&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("skrname")&"&nbsp; </div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px"">"&rspxzj("bz")&"&nbsp;</td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"
                 call editdel(rspxzj("id"),rspxzj("sscj"),"pxzj_view.asp?action=edit&id=","pxzj_view.asp?action=del&id=")
                 response.write "</div></td></tr>"
                 rspxzj.movenext
		      loop
          end if 
		  response.write "  </tr></table><br>"
		  
'��ʾ����������������ѵ		  
 sql="SELECT * from bzname where sscj="&request("sscj")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   do while not rs.eof
      sqlpxzj="SELECT * from pxzj where ssbz="&rs("id")&" and month="&request("month")&" and year="&request("year")
      set rspxzj=server.createobject("adodb.recordset")
      rspxzj.open sqlpxzj,conne,1,1
      if rspxzj.eof and rspxzj.bof then 
             response.write "<p align='center'>δ���"&ssbzh(rs("id"))&"��ѵ�ܽ�</p>" 
          else
             response.write "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
             response.write " <tr class=""title""><td colspan=13 >&nbsp;&nbsp;&nbsp;  ��λ��"&sscjh(request("sscj"))&"&nbsp;"&ssbzh(rspxzj("ssbz"))
             response.write "</td></tr>"
             response.write "<tr class=""title""><td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ʱ��</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ����ժҪ</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ����</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>Ӧ���˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ʵ���˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ϸ��˴�</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ��</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ϸ���</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>��ѵ��ʽ</div></td>"
             response.write " <td   style=""border-bottom-style: solid;border-width:1px""><div align=center>�ۼƿ�ʱ</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>�ڿ���</div></td>"
             response.write " <td   style=""border-bottom-style: solid;border-width:1px""><div align=center>��ע</div></td>"
			 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>ѡ��</div></td></tr>"
              do while not rspxzj.eof
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
				 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("month")&"."&rspxzj("day")&"</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px"">"&rspxzj("body")&"&nbsp;</td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("dx")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("numb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("sdnumb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("hgnumb")&"&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("sdnumb")/rspxzj("numb")*100&"%&nbsp;</div></td>"
             response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("hgnumb")/rspxzj("numb")*100&"%&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("xs")&"&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("ks")&"h&nbsp;</div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"&rspxzj("skrname")&"&nbsp; </div></td>"
                 response.write "<td  style=""border-bottom-style: solid;border-width:1px"">"&rspxzj("bz")&"&nbsp;</td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>"
                 call editdel(rspxzj("id"),rspxzj("sscj"),"pxzj_view.asp?action=edit&id=","pxzj_view.asp?action=del&id=")
                 response.write "</div></td></tr>"
                 rspxzj.movenext
		      loop
          end if 
		  response.write "  </tr></table><br>"
        rs.movenext
  loop
  rs.close
  set rs=nothing
  rspxzj.close
  set rspxzj=nothing
end sub


response.write "</body></html>"
'************************************
'�������¼����ʾ��Ӧ�ı༭��ɾ��
'*************************************

Call CloseConn
%>