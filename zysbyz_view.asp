<%@language=vbscript codepage=936 %>
<%
Option Explicit
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
response.write " if(document.form1.sscj.value==''){" & vbCrLf
response.write "      alert('��ѡ���������䣡');" & vbCrLf
response.write "   document.form1.sscj.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write " if(document.form1.zysb_name.value==''){" & vbCrLf
response.write "      alert('��ѡ���豸���ƣ�');" & vbCrLf
response.write "   document.form1.zysb_name.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
'������������Զ�����ҳ����ʾ��ѡ�豸��̨����λ��
	response.write "function check(){" & vbCrLf
    response.write "if(document.getElementById(""zysb_numb"").style.display==""none"")" & vbCrLf
    response.write "		document.getElementById(""zysb_numb"").style.display=""inline"";" & vbCrLf
   response.write "	var snumb=numb[document.getElementById(""zysb_name"").value]" & vbCrLf
   response.write "	document.getElementById(""zysb_numb"").innerHTML=snumb;" & vbCrLf
    response.write "	document.getElementById(""zysb_numb"").className=""ok"";" & vbCrLf
	    response.write "if(document.getElementById(""zysb_wh"").style.display==""none"")" & vbCrLf
    response.write "		document.getElementById(""zysb_wh"").style.display=""inline"";" & vbCrLf
   response.write "	var swh=wh[document.getElementById(""zysb_name"").value]" & vbCrLf
   response.write "	document.getElementById(""zysb_wh"").innerHTML=swh;" & vbCrLf
    response.write "	document.getElementById(""zysb_wh"").className=""ok"";" & vbCrLf
    response.write "	return;" & vbCrLf
	response.write "}"

response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

if request("action")="" then call main()
if request("action")="add" then call add()
if request("action")="checksaveadd" then call checksaveadd()
if request("action")="del" then call del()
if request("action")="edit" then call edit()
if request("action")="saveedit" then saveedit()
sub add()
dim ii
dim rscj,sqlcj,rsbz,sqlbz,sql,rs
   response.write"<br><br><br><form method='get' action='zysbyz_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�����Ҫ�豸��ת��</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>�豸���ƣ� </strong></td>"& vbCrLf      
    response.write"<td width='70%' class='tdbg'>"& vbCrLf
  if session("level")=0 then 
	'***************************************************������������ҳ����ʾ��ѡ�������е��豸�Թ�ѡ��
	response.write"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"" onBlur=""check()"">"& vbCrLf
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
    response.write "<select name='zysb_name' size='1'  onBlur=""check()""  >" & vbCrLf
    response.write "<option  selected>ѡ���豸</option>" & vbCrLf
    response.write "</select></td></tr>  "  & vbCrLf
    response.write "<script><!--" & vbCrLf
    response.write "var groups=document.form1.sscj.options.length" & vbCrLf
    response.write "var group=new Array(groups)" & vbCrLf
    response.write "var numb=new Array()" & vbCrLf
    response.write "var wh=new Array()" & vbCrLf
	response.write "for (i=0; i<groups; i++)" & vbCrLf
    response.write "group[i]=new Array()" & vbCrLf
    response.write "group[0][0]=new Option(""ѡ���豸"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from zysbname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,connb,1,1
        if rsbz.eof and rsbz.bof then
               response.write"group["&rscj("id")&"]["&ii&"]=new Option(""δ����豸"","""");" & vbCrLf	
		else
		do while not rsbz.eof
		   'response.write"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		   response.write"group["&rscj("levelid")&"]["&ii&"]=new Option("""&rsbz("name")&""","""&rsbz("id")&""");" & vbCrLf
		  response.write "numb["&rsbz("id")&"]="&rsbz("numb")&";" & vbCrLf
		  response.write "wh["&rsbz("id")&"]="""&rsbz("wh")&""";" & vbCrLf
		  
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
    response.write "var temp=document.form1.zysb_name" & vbCrLf
    response.write "function redirect(x){" & vbCrLf
    response.write "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    response.write "temp.options[m]=null" & vbCrLf
    response.write "for (i=0;i<group[x].length;i++){" & vbCrLf
    response.write "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    response.write "}" & vbCrLf
    response.write "temp.options[0].selected=true" & vbCrLf
    response.write "}"
	response.write "	//--></script>" & vbCrLf
'***************************************************������������ҳ����ʾ��ѡ�������е��豸�Թ�ѡ��

  else 
   
   '****************************���ݵ�¼�û�����ʾ����������豸
   response.write"<input name='sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
   response.write"<input name='sscj' type='hidden' value="&session("level")&">"& vbCrLf
   sql="SELECT * from zysbname where sscj="&session("level")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,connb,1,1
   response.write"<select name='zysb_name' size='1'  onBlur=""check()"" >"
   
   if rs.eof and rs.bof then 
   	  'response.write"<option value='0'>����</option>"
   else   
	  'response.write"<option value='0'>����</option>"
      do while not rs.eof
	     response.write"<option value='"&rs("id")&"'>"&rs("name")&"</option>"
	  rs.movenext
      loop
	  end if 
	 response.Write"</select>" 
  rs.close
  set rs=nothing
 
 
     response.write "<script><!--" & vbCrLf
    response.write "   document.form1.zysb_name.focus();" & vbCrLf  '�������Ӻ�ʹ�豸���Ŀ�Ϊ���㣬�Ա��û������ĵط�ʧȥ������Զ���ʾ̨����λ��
    response.write "var wh=new Array()" & vbCrLf
    response.write "var numb=new Array()" & vbCrLf
		sqlbz="SELECT * from zysbname where sscj="&session("level")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,connb,1,1
        if rsbz.eof and rsbz.bof then
		else
		  do while not rsbz.eof
		  response.write "numb["&rsbz("id")&"]="&rsbz("numb")&";" & vbCrLf
		  response.write "wh["&rsbz("id")&"]="""&rsbz("wh")&""";" & vbCrLf
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing
     	response.write "	//--></script>" & vbCrLf
   '****************************���ݵ�¼�û�����ʾ����������豸
    end if 
    response.write"</td></tr>  "  	 

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�豸λ�ţ�</strong></td> "
	 response.write"<td width='88%' class='tdbg'><div id=""zysb_wh"" style=""display:none"" class=""ok""></div></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�豸������</strong></td> "
	 response.write"<td width='88%' class='tdbg'><div id=""zysb_numb"" style=""display:none"" class=""ok""></div></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���¿���Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='kgby_m' value='0'></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcjxby_m' value=0></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcbyby_m' value=0></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ���¹�Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcsgby_m' value=0></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcqtby_m' value=0></td></tr> "

			 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ڣ�</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='date' type='text' value="&year(now())&"-"&month(now())&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this,date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='checksaveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub checksaveadd()
dim sqlbb,rsbb
   sqlbb="SELECT * from zysbyz where  zysb="&Request("zysb_name")&" and year="&year(Request("date"))&" and month="&month(Request("date"))
   set rsbb=server.createobject("adodb.recordset")
   rsbb.open sqlbb,connb,1,1
   if rsbb.eof and rsbb.bof then 
     call saveadd
   else
      response.write"<Script Language=Javascript>window.alert('����ӹ����豸�ı���');history.go(-1);</Script>"
   end if
end sub

sub saveadd()    
	dim sylj  '�����ۼ�Сʱ
	dim syljl  '�����ۼ�Сʱ��
	dim bylj   '�����ۼ�Сʱ��
	  dim year1,month1,day1'����\
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zysbyz" 
      rsadd.open sqladd,connb,1,3
      rsadd.addnew
     ' on error resume next
	  rsadd("sscj")=Trim(Request("sscj"))
      rsadd("zysb")=Trim(Request("zysb_name"))
      'rsadd("zysb_wh")=Trim(Request("zysb_wh"))
      'rsadd("zysb_numb")=Trim(Request("zysb_numb"))
	  year1=year(Trim(Request("date")))
	  month1=month(Trim(Request("date")))
	  if len(month1)<>2 then month1="0"&month1
      rsadd("month")=month1
	  rsadd("year")=year1
	  
	  
	  rsadd("kgby_m")=cint(request("kgby_m"))
      rsadd("tcjxby_m")=cint(Trim(request("tcjxby_m")))
      rsadd("tcbyby_m")=cint(request("tcbyby_m"))
      rsadd("tcsgby_m")=cint(request("tcsgby_m"))
      rsadd("tcqtby_m")=cint(request("tcqtby_m"))
	   'response.write request("kgby_m")&"-"&Trim(request("tcjxby_m"))&"-"&request("tcbyby_m")&"-"&request("tcsgby_m")&"-"&request("tcqtby_m")
	   
	   
	 '��ȡ��������ۼ�Сʱ�����뱾��Сʱ����ӣ��õ�  �����ۼ�Сʱ��
				dim sqlbb,rsbb          
                 if month(Request("date"))=1 then 
			        sqlbb="SELECT * from zysbyz where zysb="&Request("zysb_name")&" and year="&year(Request("date"))-1&" and month=12"
		         else
			        sqlbb="SELECT * from zysbyz where  zysb="&Request("zysb_name")&" and year="&year(Request("date"))&" and month="&month(Request("date"))-1
                 end if 
			   set rsbb=server.createobject("adodb.recordset")
               rsbb.open sqlbb,connb,1,1
               if rsbb.eof and rsbb.bof then 
                     rsadd("kglj_m")=cint(request("kgby_m"))
                     rsadd("tcjxlj_m")=cint(Trim(request("tcjxby_m")))
                     rsadd("tcbylj_m")=cint(request("tcbyby_m"))
                     rsadd("tcsglj_m")=cint(request("tcsgby_m"))
                     rsadd("tcqtlj_m")=cint(request("tcqtby_m"))
                     'bylj=cint(request("kgby_m"))
					 'sylj=0
					 'syljl=0
			   rsadd("yzllj")=cint(request("kgby_m"))/cint(request("kgby_m"))+cint(request("tcjxby_m"))+cint(request("tcbyby_m"))+cint(request("tcsgby_m"))+cint(request("tcqtby_m"))
			   else
                     rsadd("kglj_m")=cint(request("kgby_m"))+rsbb("kglj_m")
                     bylj=cint(request("kgby_m"))+rsbb("kglj_m")
					 sylj=rsbb("kglj_m")
					 syljl=rsbb("yzllj")
					 
					 rsadd("tcjxlj_m")=cint(Trim(request("tcjxby_m")))+rsbb("tcjxlj_m")
                     rsadd("tcbylj_m")=cint(request("tcbyby_m"))+rsbb("tcbylj_m")
                     rsadd("tcsglj_m")=cint(request("tcsgby_m"))+rsbb("tcsglj_m")
                     rsadd("tcqtlj_m")=cint(request("tcqtby_m"))+rsbb("tcqtlj_m")
	                         	 '��ת�ʱ��£�kgby_m/(kgby_m+tcjxby_m+tcbyby_m+tcsgby_m+tcqtby_m)
	                 '           =���¿���Сʱ��/���¼���Сʱ�����±���Сʱ������ �¹�Сʱ����������Сʱ�����¿���Сʱ
	                 '��ת���ۼ�=bylj/(sylj/syljl+kgby)
	                 rsadd("yzllj")=bylj/(sylj/syljl+cint(request("kgby_m")))
			  end if
	          rsadd("yzlby")=cint(request("kgby_m"))/(cint(request("kgby_m"))+cint(request("tcjxby_m"))+cint(request("tcbyby_m"))+cint(request("tcsgby_m"))+cint(request("tcqtby_m")))
			  rsbb.close
			  set rsbb=nothing
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  	  response.write"<Script Language=Javascript>location.href='zysbyz.asp';</Script>"

end sub



sub saveedit()    
	  	dim sylj  '�����ۼ�Сʱ
	dim syljl  '�����ۼ�Сʱ��
	dim bylj   '�����ۼ�Сʱ��

	  dim year1,month1,day1'����\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zysbyz where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connb,1,3
	  year1=year(Trim(Request("date")))
	  month1=month(Trim(Request("date")))
	  if len(month1)<>2 then month1="0"&month1
      rsedit("month")=month1
	  rsedit("year")=year1
	  
	  
	  rsedit("kgby_m")=cint(request("kgby_m"))
      rsedit("tcjxby_m")=cint(Trim(request("tcjxby_m")))
      rsedit("tcbyby_m")=cint(request("tcbyby_m"))
      rsedit("tcsgby_m")=cint(request("tcsgby_m"))
      rsedit("tcqtby_m")=cint(request("tcqtby_m"))
	   'response.write request("kgby_m")&"-"&Trim(request("tcjxby_m"))&"-"&request("tcbyby_m")&"-"&request("tcsgby_m")&"-"&request("tcqtby_m")
	   
	   
	 '��ȡ��������ۼ�Сʱ�����뱾��Сʱ����ӣ��õ�  �����ۼ�Сʱ��
				dim sqlbb,rsbb          
                 if month(Request("date"))=1 then 
			        sqlbb="SELECT * from zysbyz where zysb="&Request("zysb_name")&" and year="&year(Request("date"))-1&" and month=12"
		         else
			        sqlbb="SELECT * from zysbyz where  zysb="&Request("zysb_name")&" and year="&year(Request("date"))&" and month="&month(Request("date"))-1
                 end if 
			   set rsbb=server.createobject("adodb.recordset")
               rsbb.open sqlbb,connb,1,1
               if rsbb.eof and rsbb.bof then 
                     rsedit("kglj_m")=cint(request("kgby_m"))
                     rsedit("tcjxlj_m")=cint(Trim(request("tcjxby_m")))
                     rsedit("tcbylj_m")=cint(request("tcbyby_m"))
                     rsedit("tcsglj_m")=cint(request("tcsgby_m"))
                     rsedit("tcqtlj_m")=cint(request("tcqtby_m"))
                     'bylj=cint(request("kgby_m"))
					 'sylj=0
					 'syljl=0
			   rsedit("yzllj")=cint(request("kgby_m"))/cint(request("kgby_m"))+cint(request("tcjxby_m"))+cint(request("tcbyby_m"))+cint(request("tcsgby_m"))+cint(request("tcqtby_m"))
			   else
                     rsedit("kglj_m")=cint(request("kgby_m"))+rsbb("kglj_m")
                     bylj=cint(request("kgby_m"))+rsbb("kglj_m")
					 sylj=rsbb("kglj_m")
					 syljl=rsbb("yzllj")
					 
					 rsedit("tcjxlj_m")=cint(Trim(request("tcjxby_m")))+rsbb("tcjxlj_m")
                     rsedit("tcbylj_m")=cint(request("tcbyby_m"))+rsbb("tcbylj_m")
                     rsedit("tcsglj_m")=cint(request("tcsgby_m"))+rsbb("tcsglj_m")
                     rsedit("tcqtlj_m")=cint(request("tcqtby_m"))+rsbb("tcqtlj_m")
	                         	 '��ת�ʱ��£�kgby_m/(kgby_m+tcjxby_m+tcbyby_m+tcsgby_m+tcqtby_m)
	                 '           =���¿���Сʱ��/���¼���Сʱ�����±���Сʱ������ �¹�Сʱ����������Сʱ�����¿���Сʱ
	                 '��ת���ۼ�=bylj/(sylj/syljl+kgby)
	                 rsedit("yzllj")=bylj/(sylj/syljl+cint(request("kgby_m")))
			  end if
	          rsedit("yzlby")=cint(request("kgby_m"))/(cint(request("kgby_m"))+cint(request("tcjxby_m"))+cint(request("tcbyby_m"))+cint(request("tcsgby_m"))+cint(request("tcqtby_m")))
			  rsbb.close
			  set rsbb=nothing
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub edit()

   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zysbyz where id="&id
   rsedit.open sqledit,connb,1,1

   response.write"<br><br><br><form method='get' action='zysbyz_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�༭�豸��ת�ʱ���</strong></div></td>    </tr>"
	
	response.write"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"& vbCrLf      
    response.write"<td width='70%' class='tdbg'>"& vbCrLf
    response.write"<input  value="&sscjh(rsedit("sscj"))&" type='text' disabled='disabled' >"& vbCrLf
     response.write"<input name='sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	
	response.write"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>�豸���ƣ� </strong></td>"& vbCrLf      
    response.write"<td width='88%' class='tdbg'>"& vbCrLf
    response.write"<input value="&zysb(rsedit("zysb"))&" type='text' disabled='disabled' >"& vbCrLf
     response.write"<input name='zysb_name' type='hidden' value="&rsedit("zysb")&"></td></tr>"& vbCrLf

	response.write"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>�豸λ�ţ� </strong></td>"& vbCrLf      
    response.write"<td width='88%' class='tdbg'>"& vbCrLf
    response.write"<input value="&zysbwh(rsedit("zysb"))&" type='text' disabled='disabled' ></td></tr>"& vbCrLf

	response.write"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>�豸̨���� </strong></td>"& vbCrLf      
    response.write"<td width='88%' class='tdbg'>"& vbCrLf
    response.write"<input value="&zysbnumb(rsedit("zysb"))&" type='text' disabled='disabled' ></td></tr>"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���¿���Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='kgby_m' value="&rsedit("kgby_m")&"></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcjxby_m' value="&rsedit("tcjxby_m")&"></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcbyby_m' value="&rsedit("tcbyby_m")&"></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ���¹�Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcsgby_m' value="&rsedit("tcsgby_m")&"></td></tr> "

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͣ������Сʱ����</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='tcqtby_m' value="&rsedit("tcqtby_m")&"></td></tr> "

			 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ڣ�</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='date' type='text' value="&rsedit("year")&"-"&rsedit("month")&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this,date, 'yyyy-mm'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf


	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub

sub del()
 dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from zysbyz where id="&request("id")
  rsdel.open sqldel,connb,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub main()
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

dim xh
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>��Ҫ�豸��ת��</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""zysbyz.asp"">��Ҫ�豸��ת����ҳ</a>&nbsp;|&nbsp;<a href=""zysbyz_view.asp?action=add"">�����ת��</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=zysbname"">�豸����</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=addsb"">����豸</a>"& vbCrLf
response.write " </td> </tr>"& vbCrLf
response.write "</table>"& vbCrLf
dim rszysbyz,sqlzysbyz,rs,sql
   response.write"<div align=center><strong>"&sscjh(request("sscj"))&request("year")&"��"&request("month")&"�·���Ҫ�豸��ת��</strong></div>"
   '��ʾ���伶����ѵ�ƻ�
      sqlzysbyz="SELECT * from zysbyz where sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rszysbyz=server.createobject("adodb.recordset")
      rszysbyz.open sqlzysbyz,connb,1,1
      if rszysbyz.eof and rszysbyz.bof then 
             response.write "<p align='center'>δ�������</p>" 
          else
     response.write "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  width=""100%"">"
     response.write "<tr class=""title"" >"
  response.write "<td rowspan=4  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">���</div></td>"& vbCrLf
  response.write "<td rowspan=3  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">�豸λ��</div></td>"& vbCrLf
 response.write " <td rowspan=3 style=""border-bottom-style: solid;border-width:1px""  ><div align=""center"">�豸����</div></td>"& vbCrLf
 response.write " <td rowspan=3 style=""border-bottom-style: solid;border-width:1px""  ><div align=""center"">�豸̨��</div></td>"& vbCrLf
 response.write " <td colspan=2 rowspan=2 style=""border-bottom-style: solid;border-width:1px""  ><div align=""center"">����Сʱ��</div></td>"& vbCrLf
 response.write " <td colspan=8  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">ͣ��Сʱ��</div></td>"& vbCrLf
response.write "  <td colspan=2  style=""border-bottom-style: solid;border-width:1px"" rowspan=2 ><div align=""center"">��ת��</div></td>"& vbCrLf
response.write "  <td colspan=2  style=""border-bottom-style: solid;border-width:1px"" rowspan=2 ><div align=""center"">������</div></td>"& vbCrLf
response.write "  <td rowspan=4 style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">ѡ��</div></td>"& vbCrLf
response.write " </tr>"& vbCrLf
     response.write "<tr class=""title"" >"
response.write "  <td colspan=2 style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td colspan=2 style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td colspan=2 style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">�¹�</div></td>"& vbCrLf
response.write "  <td colspan=2 style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">����</div></td>"& vbCrLf
response.write " </tr>"& vbCrLf
     response.write "<tr class=""title"" >"
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "<td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">����</div></td>"& vbCrLf
response.write "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">�ۼ�</div></td>"& vbCrLf
response.write " </tr>"& vbCrLf
response.write "<tr class=""title"" >"
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">1</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">2</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">3</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">4</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">5</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">6</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">7</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">8</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">9</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">10</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">11</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">12</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">13</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">14</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">15</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">16</div></td>"& vbCrLf
response.write "   <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">17</div></td>"& vbCrLf
response.write "  </tr> "& vbCrLf            
do while not rszysbyz.eof
 xh=xh+1
  response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
 response.write "<td><div align=""center"">"&xh&"</div></td>"
 response.write "<td>"&zysbwh(rszysbyz("zysb"))&"</td>"
 response.write "<td>"&zysb(rszysbyz("zysb"))&"</td> "& vbCrLf     
  response.write "<td><div align=""center"">"&zysbnumb(rszysbyz("zysb"))&"</div></td> "& vbCrLf     
  response.write "<td>"&rszysbyz("kgby_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("kglj_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcjxby_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcjxlj_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcbyby_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcbylj_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcsgby_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcsglj_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcqtby_m")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("tcqtlj_m")&"</td> "& vbCrLf     
  response.write "<td>"&left(rszysbyz("yzlby")*100,6)&"</td> "& vbCrLf     
  response.write "<td>"&left(rszysbyz("yzllj")*100,6)&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("cllby")&"</td> "& vbCrLf     
  response.write "<td>"&rszysbyz("clllj")&"</td> "& vbCrLf     
  response.write "<td><div align=center>"
	call editdel(rszysbyz("id"),rszysbyz("sscj"),"zysbyz_view.asp?action=edit&id=","zysbyz_view.asp?action=del&id=")
				
 response.write "</div> </tr> "& vbCrLf     
		   rszysbyz.movenext
			  loop
          end if 
		  response.write " </table>"
		  
  rszysbyz.close
  set rszysbyz=nothing
end sub


response.write "</body></html>"
'************************************
'�������¼����ʾ��Ӧ�ı༭��ɾ��
'*************************************

function zysb(zysbid)
  dim sqlbz,rsbz
 sqlbz="SELECT * from zysbname where id="&zysbid
 set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,connb,1,1
        if rsbz.eof and rsbz.bof then
               'zysbname1="����"
		else
               zysb=rsbz("name")
	    end if 
		rsbz.close
	    set rsbz=nothing
end function

function zysbwh(zysbid)
  dim sqlbz,rsbz
 sqlbz="SELECT * from zysbname where id="&zysbid
 set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,connb,1,1
        if rsbz.eof and rsbz.bof then
               'zysbname1="����"
		else
               zysbwh=rsbz("wh")
	    end if 
		rsbz.close
	    set rsbz=nothing
end function

function zysbnumb(zysbid)
  dim sqlbz,rsbz
 sqlbz="SELECT * from zysbname where id="&zysbid
 set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,connb,1,1
        if rsbz.eof and rsbz.bof then
               'zysbname1="����"
		else
               zysbnumb=rsbz("numb")
	    end if 
		rsbz.close
	    set rsbz=nothing
end function

Call CloseConn
%>