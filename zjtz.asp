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
'on error resume next
url=geturl
dim keys,sscjid,title,classid
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
classid=trim(request("classid")) 
dim url,lb,brxx,sql,rs,record,pgsz,total,page,start,rowcount,ii
dim rsadd,sqladd,id,rsdel,sqldel,rsedit,sqledit
dim sqlscdate,rsscdate'�ϴ��ܼ�ʱ��
dim zjmonth '�ܼ��·�
Dwt.Out "<html>"& vbCrLf
Dwt.Out "<head>" & vbCrLf
Dwt.Out "<title>��Ϣ����ϵͳ�ܼ�̨�˹���ҳ</title>"& vbCrLf
Dwt.Out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.Out "<SCRIPT language=javascript>" & vbCrLf
Dwt.Out "function checkadd(){" & vbCrLf
Dwt.Out "if(document.form1.sscj.value==''){" & vbCrLf
Dwt.Out "alert('��ѡ���������䣡');" & vbCrLf
Dwt.Out "document.form1.sscj.focus();" & vbCrLf
Dwt.Out "return false;" & vbCrLf
Dwt.Out "}" & vbCrLf

Dwt.Out "if(document.form1.zjtz_wh.value==''){" & vbCrLf
Dwt.Out "alert('λ�Ų���Ϊ�գ�');" & vbCrLf
Dwt.Out "document.form1.zjtz_wh.focus();" & vbCrLf
Dwt.Out "return false;" & vbCrLf
Dwt.Out "}" & vbCrLf
Dwt.Out "}" & vbCrLf
Dwt.Out "</SCRIPT>" & vbCrLf
Dwt.Out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out"<script language=javascript src='/js/popselectdate.js'></script>"
Dwt.Out "</head>"& vbCrLf
Dwt.Out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
action=request("action")
select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "editd"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editd
	'call edit
  case "saveeditd"
    call saveeditd
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
  case "history"
    call history
  case "editinfo"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editinfo
	'call edit
  case "saveeditinfo"
    call saveeditinfo
  case "delinfo"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call delinfo
	
end select	





'
'select case action
'  case "add"
'    if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
'  case "saveadd"
'    call saveadd
'  case "editd"
'	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editd
'  case "saveedit"
'    call saveedit
'  case "del"
'    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
'  case ""
'	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
'end select	



Sub add()
	Dwt.Out"<br><br><br><form method='post' action='zjtz.asp' name='form1' onSubmit='javascript:return checkadd();'>"
	Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.Out"<tr class='title'><td height='22' colspan='2'>"
	Dwt.Out"<Div align='center'><strong>����ܼ�</Div></td>    </tr>"
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
	Dwt.Out"<td width='80%' class='tdbg'>"
	if session("level")=0 then 
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select>"  	 & vbCrLf
    dwt.out "<select name='ssbz' size='1' >" & vbCrLf
    dwt.out "<option  selected>ѡ��������</option>" & vbCrLf
    dwt.out "</select></td></tr>  "  & vbCrLf
    dwt.out "<script><!--" & vbCrLf
    dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
    dwt.out "var group=new Array(groups)" & vbCrLf
    dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    dwt.out "group[i]=new Array()" & vbCrLf
    dwt.out "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   dwt.out "group["&rscj("levelid")&"][0]=new Option(""����"",""0"");" & vbCrLf
		else
		   dwt.out"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		do while not rsbz.eof
		   dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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




    dwt.out "var temp=document.form1.ssbz" & vbCrLf
    dwt.out "function redirect(x){" & vbCrLf
    dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    dwt.out "temp.options[m]=null" & vbCrLf
    dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    dwt.out "}" & vbCrLf
    dwt.out "temp.options[0].selected=true" & vbCrLf
    dwt.out "}//--></script>" & vbCrLf



  else 	 
   dwt.out"<input name='sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   dwt.out"<select name='ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  dwt.out"<option value='0'>����</option>"
   else   
	  dwt.out"<option value='0'>����</option>"
      do while not rs.eof
	     dwt.out"<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	  end if 
	 dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
	Dwt.Out"</td></tr>"& vbCrLf
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	Dwt.Out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_wh' type='text'></td>    </tr>   "
	 
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
	Dwt.Out"<td><select name='zjtz_lx' size='1'>"
	dim sqlname,rsname
	sqlname="SELECT * from class "
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    do while not rsname.eof
		Dwt.Out "<option value='"&rsname("id")&"'>"&rsname("name")&"</option>"
		rsname.movenext
	loop
	end if 
	rsname.close
	set rsname=nothing
    Dwt.Out"</select></td></tr>"
	 
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ʽ��</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_glfs' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������ţ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ccbh' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ggxh' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>���ȵȼ���</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_jddj' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_clfw' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict ("zjtz_jdzq","��������",onnumb)
	dwt.out "</td></tr>"
	
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ϴ��ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
	Dwt.out "<input type='checkbox' name='isdx' onclick='zjtz_dxyear.disabled=!checked;zjtz_date.disabled=checked;'/>�Ƿ�������ܼ�"
	Dwt.out "<br/><select name='zjtz_dxyear' disabled='disabled'/>"
	for  i=year(now())-5 to year(now())+5
         Dwt.out "<option value="&i
		 if i=year(now()) then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>�����ܼ����"
    Dwt.out"<br/><input name='zjtz_date'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'/>�ճ��ܼ�����"
	
	
	Dwt.Out"<br/>����ݼ������ں��´μ���ʱ���������һ��ģ����ϴ��ܼ�����</td>    </tr>   "
		
    'Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>���������</strong></td>"
	'Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_info' type='text'></td>    </tr>   "

	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    Dwt.Out"<td width='80%' class='tdbg'><input type='text' name='zjtz_bz'></td></tr>  "   

	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='Submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.Out"</table></form>"
end Sub	

Sub saveadd()    
	  'on error resume next
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zjtz" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("sscj"))
      rsadd("ssbz")=Trim(Request("ssbz"))
	  rsadd("wh")=request("zjtz_wh")
      rsadd("ggxh")=request("zjtz_ggxh")
      rsadd("jddj")=request("zjtz_jddj")
      rsadd("clfw")=request("zjtz_clfw")
      rsadd("jdzq")=cint(request("zjtz_jdzq"))
      rsadd("glfs")=request("zjtz_glfs")
      rsadd("ccbh")=request("zjtz_ccbh")
	  rsadd("class")=cint(request("zjtz_lx"))
      if request("isdx")="on" then 
	     rsadd("dxzjyear")=request("zjtz_dxyear")
	     rsadd("isdx")=true
		' message request("isdx")&request("zjtz_dxyear")
	  else
	     rsadd("sczjdate")=request("zjtz_date")
	     rsadd("isdx")=false
	  end if 
	  
	  rsadd("bz")=request("zjtz_bz")
	  rsadd.update
      
'	  dim rsinfo,sqlinfo
'	        set rsinfo=server.createobject("adodb.recordset")
'      sqlinfo="select * from zjinfo" 
'      rsinfo.open sqlinfo,connzj,1,3
'      rsinfo.addnew
'      rsinfo("zjtzid")=rsadd("id")
'      rsinfo("zjyear")=cint(Request("zjtz_year"))
'	  rsinfo("zjmonth")=request("zjtz_month")
'      rsinfo("zjday")=request("zjtz_day")
'      rsinfo("zjinfo")=request("zjtz_info")
'	  rsinfo.update
'	  rsinfo.close
'      set rsinfo=nothing
'
'	  
	  rsadd.close
      set rsadd=nothing
	  
	  
	  
	  Dwt.Out"<Script Language=Javascript>location.href='zjtz.asp';</Script>"
end Sub

Sub saveeditd()    
      'on error resume next
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zjtz where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
      rsedit("sscj")=Trim(Request("sscj"))
      rsedit("ssbz")=Trim(Request("ssbz"))
	  rsedit("wh")=request("zjtz_wh")
      rsedit("ggxh")=request("zjtz_ggxh")
	  rsedit("glfs")=request("zjtz_glfs")
      rsedit("ccbh")=request("zjtz_ccbh")
      rsedit("clfw")=request("zjtz_clfw")
      rsedit("jddj")=request("zjtz_jddj")
      'rsedit("sczjdate")=request("zjtz_sczjdate")
      rsedit("jdzq")=cint(request("zjtz_jdzq"))
	  rsedit("class")=cint(request("zjtz_lx"))
      rsedit("bz")=request("zjtz_bz")
      if request("isdx")="on" then 
	     rsedit("dxzjyear")=request("zjtz_dxyear")
	     rsedit("isdx")=true
		' message request("isdx")&request("zjtz_dxyear")
	  else
	     rsedit("sczjdate")=request("zjtz_date")
	  	 rsedit("isdx")=false

	  end if 

	  
	  
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.Out"<Script Language=Javascript>history.go(-2)</Script>"
end Sub

Sub del()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from zjtz where id="&id
  rsdel.open sqldel,connzj,1,3
  set rsdel=nothing  
  
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete from zjinfo where zjtzid="&id
  rsdel.open sqldel,connzj,1,3
  'set rsdel=nothing  
  
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end Sub
Sub delinfo()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from zjinfo where id="&id
  rsdel.open sqldel,connzj,1,3
  set rsdel=nothing  
  
  
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end Sub


Sub editd()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zjtz where id="&id
   rsedit.open sqledit,connzj,1,1
   Dwt.Out"<br><br><br><form method='post' action='zjtz.asp' name='form1' onSubmit='javascript:return checkadd();'>"& vbCrLf
   Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   Dwt.Out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.Out"<Div align='center'><strong>�༭�ܼ�</Div></td>    </tr>"& vbCrLf
     Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     Dwt.Out"<td width='80%' class='tdbg'>"& vbCrLf
     Dwt.Out"<input name='sscj' type='hidden' value="&rsedit("sscj")&">"& vbCrLf

	dim sqlbz,rsbz
	sqlbz="SELECT * from bzname where sscj="&rsedit("sscj")
   	set rsbz=server.createobject("adodb.recordset")
   	rsbz.open sqlbz,conn,1,1
   	Dwt.Out"<select name='ssbz' size='1'>"
   	if rsbz.eof and rsbz.bof then 
   		  Dwt.Out"<option value='0'>δ��Ӱ���</option>"& vbCrLf
   	else   
		  'Dwt.Out"<option value='0'>����</option>"
   	   do while not rsbz.eof
		     Dwt.Out"<option value='"&rsbz("id")&"'"
			 if rsedit("ssbz")=rsbz("id") then Dwt.Out " selected"
			 Dwt.Out">"&rsbz("bzname")&"</option>"& vbCrLf
		  rsbz.movenext
   	   loop
	end if 
	 Dwt.Out"</select>" & vbCrLf
 	 rsbz.close
 	 set rsbz=nothing
	 Dwt.Out"</td></tr>"& vbCrLf


	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 Dwt.Out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_wh' type='text' value="""&rsedit("wh")&"""></td>    </tr>   "
	 
	 
	 Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
	Dwt.Out"<td><select name='zjtz_lx' size='1'>"
	dim sqlname,rsname
	sqlname="SELECT * from class "
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    do while not rsname.eof
		Dwt.Out "<option value='"&rsname("id")&"'"
        if rsedit("class")=rsname("id") then Dwt.Out "selected"
		Dwt.Out ">"&rsname("name")&"</option>"
		rsname.movenext
	loop
	end if 
	rsname.close
	set rsname=nothing
    Dwt.Out"</select></td></tr>"
	 
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ʽ��</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_glfs' type='text' value="&rsedit("glfs")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������ţ�</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ccbh' type='text' value="&rsedit("ccbh")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ggxh' type='text' value="&rsedit("ggxh")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>���ȵȼ���</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_jddj' type='text' value="&rsedit("jddj")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_clfw' type='text' value="&rsedit("clfw")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict2 ("zjtz_jdzq","��������",onnumb,rsedit("jdzq"))
    dwt.out "</td></tr>"
'					Dwt.Out "      <td  class='x-td'><Div align=""center"">"
'					dwt.out dispalydatadict("��������",rs("jdzq"))
'					dwt.out"&nbsp;</Div></td>" & vbCrLf
'	 Dwt.Out"<td width='80%' class='tdbg'><select name='zjtz_jdzq' size='1'>"
'      Dwt.Out "<option value='12'"
'      if rsedit("jdzq")=12 then Dwt.Out "selected"
'	  Dwt.Out ">12����</option>"
'      Dwt.Out "<option value='24'"
'      if rsedit("jdzq")=24 then Dwt.Out "selected"
'	  Dwt.Out ">24����</option>"
'      Dwt.Out "<option value='36'"
'      if rsedit("jdzq")=36 then Dwt.Out "selected"
'	  Dwt.Out">36����</option>"
'      Dwt.Out "<option value='0'"
'      if rsedit("jdzq")=0 then Dwt.Out "selected"
'	  Dwt.Out">ͣ��</option>"
'      Dwt.Out "<option value='1'"
'      if rsedit("jdzq")=1 then Dwt.Out "selected"
'      Dwt.Out">���ܼ�</option>"
'       Dwt.Out "</select></td></tr>"
'    
	    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ϴ��ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
	Dwt.out "<input type='checkbox' name='isdx' "
	if rsedit("isdx") then dwt.out "checked"
	dwt.out " onclick='zjtz_dxyear.disabled=!checked;zjtz_date.disabled=checked;'/>�Ƿ�������ܼ�"
	Dwt.out "<br/><select name='zjtz_dxyear'"
	if rsedit("isdx")=false then dwt.out " disabled='disabled'"
	dwt.out ">"
	for  i=year(now())-5 to year(now())+5
         Dwt.out "<option value="&i
		 if i=rsedit("dxzjyear") then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>�����ܼ����"
    Dwt.out"<br/><input name='zjtz_date' "
	if rsedit("isdx") then dwt.out "disabled='disabled'"
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("sczjdate")&"'/>�ճ��ܼ�����"
    
	
	'Dwt.Out"<tr><td width='20%' align='right' class='tdbg'><strong>�ϴ��ܼ����ڣ�</strong></td><td><input name='zjtz_sczjdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("sczjdate")&"'></td></tr>"
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    Dwt.Out"<td width='80%' class='tdbg'><input type='text' name='zjtz_bz' value="&rsedit("bz")&"></td></tr>  "   
	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"& vbCrLf
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveeditd'> <input type='hidden' name='id' value='"&id&"'>      <input  type='Submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"& vbCrLf
	Dwt.Out"</table></form>"& vbCrLf
	       rsedit.close
       set rsedit=nothing
end Sub
Sub editinfo()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zjinfo where id="&id
   rsedit.open sqledit,connzj,1,1
   Dwt.Out"<br><br><br><form method='post' action='zjtz.asp' name='form1' >"& vbCrLf
   Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   Dwt.Out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.Out"<Div align='center'><strong>�༭�ܼ���ʷ</strong></Div></td>    </tr>"& vbCrLf
'	if rsedit("isdx") then 
'		Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
'		Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_date' onClick='new Calendar(0).show(this)' value='"&rsedit("zjdate")&"'/></td>    </tr>   "
'	else 
'		Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
'		Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_date' onClick='new Calendar(0).show(this)' value='"&rsedit("zjdate")&"'/></td>    </tr>   "
'    end if 
	
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
	Dwt.out "<input type='checkbox' name='isdx' "
	if rsedit("isdx") then dwt.out "checked "
	dwt.out "onclick='zjtz_dxyear.disabled=!checked;zjtz_date.disabled=checked;'/>�Ƿ��Ǵ���"
	Dwt.out "<br/><select name='zjtz_dxyear'"
	if rsedit("isdx")=false then dwt.out " disabled='disabled'"
	dwt.out ">" 
	for  i=year(now())-5 to year(now())+5
         Dwt.out "<option value="&i
		 if i=rsedit("dxzjyear") then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>�����ܼ����"
    Dwt.out"<br/><input name='zjtz_date' "
	if rsedit("isdx") then dwt.out "disabled='disabled'"
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&request("zjdate")&"'/>�ճ��ܼ�����"		
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ�����</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjinfo' type='text' value="&rsedit("zjinfo")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��ע��</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='bz' type='text' value="&rsedit("bz")&"></td>    </tr>   "
	
	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"& vbCrLf
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveeditinfo'> <input type='hidden' name='id' value='"&id&"'>      <input  type='Submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"& vbCrLf
	Dwt.Out"</table></form>"& vbCrLf
	       rsedit.close
       set rsedit=nothing
end Sub
sub saveeditinfo()
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zjinfo where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
      if request("isdx")="on" then 
	     rsedit("dxzjyear")=request("zjtz_dxyear")
	     rsedit("isdx")=true
		 zjyear=request("zjtz_dxyear")
		 zjmonth=0
		' message request("isdx")&request("zjtz_dxyear")
	  else
	     rsedit("zjdate")=request("zjtz_date")
	     rsedit("isdx")=false
		 zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
	  end if 
      zjtzid=rsedit("zjtzid")
	  rsedit("bz")=request("bz")
      rsedit("zjinfo")=request("zjinfo")
	  rsedit.update
      set rsedit=nothing
	  
	  	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zjtz where id="&zjtzid
      rsedit.open sqledit,connzj,1,3
      if request("isdx")="on" then 
	     rsedit("dxzjyear")=request("zjtz_dxyear")
	     rsedit("isdx")=true
		' message request("isdx")&request("zjtz_dxyear")
	  else
	     rsedit("sczjdate")=request("zjtz_date")
	  	 rsedit("isdx")=false
	  end if 
	  
	  rsedit.update
      rsedit.close
      set rsedit=nothing
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end sub
Sub history()

    sql="SELECT * from zjtz where id="&request("id")&" ORDER BY id DESC"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzj,1,1
    if rs.eof and rs.bof then 
        Dwt.Out "<p align='center'>δ�ҵ�����</p>" 
    else
		Dwt.Out "<Div style='left:6px;'>"& vbCrLf
		Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
		Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&rs("wh")&"  �ܼ���ʷ</span>"& vbCrLf
		Dwt.Out "     </Div>"& vbCrLf
       
		'Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>����ʽ</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�ͺ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>���ȵȼ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))&"</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("wh")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("glfs")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("ccbh")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("jddj")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
         Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("jdzq")&"&nbsp;</Div></td>" & vbCrLf
        Dwt.Out "</tr></table>" & vbCrLf
    'Dwt.Out "</Div>"
   ' Dwt.Out "</Div>"
	  sscjid=rs("sscj")
	end if
	
	
    rs.close
    set rs=nothing
	
	sqlscdate="SELECT * from zjinfo where zjtzid="&request("id")&" ORDER BY id DESC"
    set rsscdate=server.createobject("adodb.recordset")
    rsscdate.open sqlscdate,connzj,1,1
    if rsscdate.eof and rsscdate.bof then 
        message("û����ǰ���ܼ��¼")
    else
         record=rsscdate.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsscdate.PageSize = Cint(PgSz) 
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
           rsscdate.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsscdate.PageSize
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>�ܼ�����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
		   do while not rsscdate.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&xh_id&"</Div></td>" & vbCrLf
        'zjmonth=month(rsscdate("zjdate"))
		'if zjmonth=0 then zjmonth="����"
                if rsscdate("isdx") then
                      zjdate=rsscdate("dxzjyear")&"-����"
                else
                      zjdate=rsscdate("zjdate")
                end if 
		Dwt.Out "      <td  class='x-td'>"&zjdate&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("zjinfo")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("bz")&"&nbsp;</td>" & vbCrLf
		'dwt.out session("levelclass")&"-"&sscjid
		if session("levelclass")=sscjid or session("levelclass")=0 then 
			Dwt.Out "<td  class='x-td'><a href=zjtz.asp?action=editinfo&id="&rsscdate("id")&">�༭</a>&nbsp;"
			Dwt.Out "<a href=zjtz.asp?action=delinfo&id="&rsscdate("id")&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a></td>"
		 else
			Dwt.Out "&nbsp;"
		 end if 
 
			 RowCount=RowCount-1
          rsscdate.movenext
          loop
        Dwt.Out "</table>" & vbCrLf
       url="zjtz.asp?action=history&id="&request("id")
	   call showpage(page,url,total,record,PgSz)
	   Dwt.Out "</Div>"
	   end if
	   Dwt.Out "</Div>"
	          rsscdate.close
	         Dwt.Out "<br><table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td>" 
			Dwt.Out "<input name='Cancel' type='button' id='Cancel' value=' ��  �� ' onClick="";history.back()"" style='cursor:hand;'></td></tr></table>"

end Sub







Sub main()
	'dim sql,rsjxjl,title
	sql="SELECT * from zjtz"
	if keys<>"" then 
		sql=sql&" where wh like '%" &keys& "%' "
		title="-���� "&keys
	end if 
	if sscjid<>"" then
	    if classid<>"" then
		sql=sql&" where sscj="&sscjid&"and class="&classid
		title="-"&sscjh(sscjid)
		else
		sql=sql&" where sscj="&sscjid
		title="-����"&sscjh(sscjid)
		end if
	else 
	    if classid<>"" then  
	    sql=sql&" where class="&classid 
	    title="-111" 
		end if
	end if 
	
	sql=sql&" ORDER BY sscj aSC "
	
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>�ܼ�̨��"&title&"</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf

	'Dwt.Out "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
'	Dwt.Out "<tr class='topbg'>"& vbCrLf
'	Dwt.Out "<td height='22' colspan='2' align='center'><strong>���ڼ춨̨�˹���ҳ</strong></td>"& vbCrLf
'	Dwt.Out "</tr>  "& vbCrLf
'	Dwt.Out "<tr class='tdbg'>"& vbCrLf
'	Dwt.Out "<td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
'	Dwt.Out "<td height='30'><a href=""zjtz.asp"">���ڼ춨̨����ҳ</a>&nbsp;|&nbsp;<a href=""zjtz.asp?action=add"">����ܼ�</a></td>"& vbCrLf
'	Dwt.Out "</tr>"& vbCrLf
'	Dwt.Out "</table>"& vbCrLf

call search()
'dim v1,v2,v3,yzj,wzj
'
'wzj=0
'yzj=0
'   dim sql,rs,rsscdate,sqlscdate,zjmonth,zjmonthname
'   sql="SELECT * from zjtz where sscj=1 "
'   set rs=server.createobject("adodb.recordset")
'   rs.open sql,connzj,1,1
'   if rs.eof and rs.bof then 
'      dim text
'   else
'      do while not rs.eof
'          dim jdzq  '�춨�����ж�
'		  dim jdyear '�춨���ڻ���Ϊ��
'		  jdzq=rs("jdzq")
'		      if jdzq=1 then 
'			    else
'				  jdyear=jdzq/12
'		          sqlscdate="SELECT * from zjinfo where zjtzid="&rs("id")
'				  'zjyear="&request("zjyear")-jdyear&" and zjmonth="&request("zjmonth")
'                  set rsscdate=server.createobject("adodb.recordset")
'                  rsscdate.open sqlscdate,connzj,1,1
'                  if rsscdate.eof and rsscdate.bof then 
'                       'Dwt.Out "<td><Div align=center>δ�ҵ�����,�����ܼ�̨������Ӵ˱�ĳ����ܼ�����</Div></td></tr>" 
'                  else
'					   if year(rsscdate("zjdate"))=year(now()) and month(rsscdate("zjdate"))=month(now())  then
'					       yzj=yzj+1
'                       else 
'						  if year(rsscdate("zjdate"))=year(now())-jdyear and month(rsscdate("zjdate"))=month(now())  then
'							wzj=wzj+1
'						  end if 
'                      end if 
'				 end if 
'			    rsscdate.close
'			end if	 	  
'       rs.movenext
'     loop
'  end if 	 
'   rs.close
'   set rs=nothing
'
'
'v1= "�������ܼ�"&yzj&"δ�ܼ�"&wzj
'
'Dwt.Out v1



	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
		Dwt.Out "<p align='center'>δ�������</p>" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ʽ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ͺ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>���ȵȼ�</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ϴμ���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�´μ���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
		'Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		record=rs.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
				   PgSz=Trim(Request("PgSz"))
			   end if 
			   rs.PageSize = Cint(PgSz) 
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
			   rs.absolutePage = page
			   start=PgSz*Page-PgSz+1
			   rowCount = rs.PageSize
		do while not rs.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
					Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))

call edit(rs("id"),rs("sscj"))
DWT.OUT "</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&searchH(uCase(rs("wh")),keys)&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("glfs")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("ccbh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("jddj")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
					
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"
					dwt.out dispalydatadict("��������",rs("jdzq"))
					dwt.out"&nbsp;</Div></td>" & vbCrLf
						
					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					jdzq=rs("jdzq")/12
					
	'				if jdzq=0 then 
	'				  Dwt.Out "<td><font color=#ff0000><Div align=center>ͣ��</Div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
	'				  Dwt.Out "      <td  class='x-td'><Div align=""center"">&nbsp;</Div></td>" & vbCrLf
	'				else
	'				  if jdzq=1 then 
	'    				  Dwt.Out "<td><font color=#ff0000><Div align=center>���ܼ�</Div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
	'				  	  Dwt.Out "      <td  class='x-td'><Div align=""center"">&nbsp;</Div></td>" & vbCrLf
	'				  else
	'				    Dwt.Out "      <td  class='x-td'><Div align=""center"">"&jdzq&"&nbsp;</Div></td>" & vbCrLf
	'					jdyear=jdzq/12
	'					'sqlscdate="SELECT * from zjinfo where zjtzid="&rs("id")&" ORDER BY id DESC"
	'                	set rsscdate=server.createobject("adodb.recordset")
	'                	rsscdate.open sqlscdate,connzj,1,1
	'                	if rsscdate.eof and rsscdate.bof then 
	'                	       Dwt.Out "<td>δ�ҵ�����</td><td>δ�ҵ�����</td><td>δ֪</td>" 
	'                	else
							   'zjmonth=month(rsscdate("zjdate"))
	'						   if zjmonth=0 then zjmonth="����"
			'�ϴ��ܼ�����
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"				   
if  rs("jdzq")<>1 then				
if rs("isdx") then 
			     Dwt.out rs("dxzjyear")&"-"&"����"
			else
			     Dwt.out rs("sczjdate")
			end if 	 	 
end if 
			Dwt.out "</Div></td>" & vbCrLf
			 'Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rsscdate("zjinfo")&"</Div></td>" & vbCrLf
			
			
			'�´��ܼ�����
			Dwt.Out "<td  class='x-td'><Div align=""center"">"
                if  rs("jdzq")<>1 then			
                        if rs("isdx") then 
			     Dwt.out rs("dxzjyear")+jdzq&"-"&"����"
			else
			     'Dwt.out year(rs("sczjdate"))+jdzq&"-"&month(rs("sczjdate"))
                             Dwt.out dateadd("m",rs("jdzq"),rs("sczjdate"))
			end if 	 	 
		end if 	

Dwt.out "</Div></td>" & vbCrLf
	'               		end if 
	'					rsscdate.close
	'				  end if 
	'				end if   
					Dwt.Out "      <td  class='x-td'>"&rs("bz")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
					'call edit(rs("id"),rs("sscj"))
					Dwt.Out "</Div></td></tr>" & vbCrLf
					 RowCount=RowCount-1
			  rs.movenext
			  loop
			Dwt.Out "</table>" & vbCrLf
		   if sscjid<>"" or keys<>"" then 
		       call showpage(page,url,total,record,PgSz)
			else
		       call showpage1(page,url,total,record,PgSz)
           end if 
		   Dwt.Out "</Div>"
		   end if
		   Dwt.Out "</Div>"		   
		   rs.close
		   set rs=nothing
end Sub
Dwt.Out "</body></html>"

'���ڷ���������ʾ
Function zjclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from class where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	'do while not rsname.eof
	else
	    zjclass=rsname("name")
		'rsname.movenext
	'loop
	end if 
	rsname.close
	set rsname=nothing
end Function

Sub edit(id,sscj)
    Dwt.Out " <a href=zjtz.asp?action=history&id="&id&">ʷ</a>&nbsp;"
if session("levelclass")=sscj or session("levelclass")=0 then 
    Dwt.Out "<a href=zjtz.asp?action=editd&id="&id&">��</a>&nbsp;"
	Dwt.Out "<a href=zjtz.asp?action=del&id="&id&" onClick=""return confirm('�˲�����ɾ���ñ����е��ܼ��¼��ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
 else
    Dwt.Out "&nbsp;"
 end if 
end Sub




Sub search()
	dim sqlcj,rscj
    Dwt.Out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.Out "<form method='Get' name='SearchForm' action='zjtz.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.Out "<a href=""zjtz.asp?action=add"">����ܼ�</a>"
	'Dwt.Out "&nbsp;&nbsp;<a href='lsda.asp?update=update'>�����������</a>"
	Dwt.Out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if keys<>"" then 
	 Dwt.Out "value='"&keys&"'"
    	Dwt.Out ">" & vbCrLf
    else
	 Dwt.Out "value='����������λ��'"
	 	Dwt.Out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	Dwt.Out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
	'Dwt.Out "  <input type='hidden' name='search' value='keys'>" & vbCrLf
	
	Dwt.Out "<select id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.Out "	       <option value=''>��������ת����</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			Dwt.Out"<option value='zjtz.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then Dwt.Out" selected"

			Dwt.Out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		Dwt.Out "     </select>	" & vbCrLf

	Dwt.Out "<select id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.Out "	       <option value=''>��������ת����</option>" & vbCrLf
	sqlcj="SELECT * from class "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,connzj,1,1
		do while not rscj.eof
			Dwt.Out"<option value='zjtz.asp?sscj="&cint(request("sscj"))&"&classid="&rscj("ID")&"'"
			if cint(request("classid"))=rscj("ID") then Dwt.Out" selected"

			Dwt.Out ">"&rscj("name")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		Dwt.Out "     </select>	" & vbCrLf

	
	
    Dwt.Out "</form></Div></Div>" & vbCrLf
end Sub





Call Closeconn
%>