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
'on error resume next

dim url,lb,brxx,sqlfdbw,rsfdbw,record,pgsz,total,page,start,rowcount,ii
dim rsadd,sqladd,id,rsdel,sqldel,rsedit,sqledit
'url=geturl
dim keys,sscjid,ssghid
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>��Ϣ����ϵͳ�������¹���ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out "if(document.form1.fdbw_sscj.value==''){" & vbCrLf
dwt.out "alert('��ѡ���������䣡');" & vbCrLf
dwt.out "document.form1.fdbw_sscj.focus();" & vbCrLf
dwt.out "return false;" & vbCrLf
dwt.out "}" & vbCrLf

dwt.out "if(document.form1.fdbw_wh.value==''){" & vbCrLf
dwt.out "alert('λ�Ų���Ϊ�գ�');" & vbCrLf
dwt.out "document.form1.fdbw_wh.focus();" & vbCrLf
dwt.out "return false;" & vbCrLf
dwt.out "}" & vbCrLf

dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

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

sub add()
dim rscj,sqlcj
   dwt.out"<br><br><br><form method='post' action='fdbw.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table  border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>��ӷ������±�</strong></div></td>    </tr>"
   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
   dwt.out"<td  class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='fdbw_sscj' size='1'>"
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
    dwt.out"<input name='fdbw_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    dwt.out"<input name='fdbw_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
 end if 
	 	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
     dwt.out"<td  class='tdbg'>"
	 	dwt.out"<select name='fdbw_gh' size='1' >"
     call formgh (0,session("levelclass"))
    dwt.out"</select> ��û����Ҫ�Ĺ���װ��,����ϵ�����������Ӧװ�ù�������</td></tr>"

	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'>"
	 dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 dwt.out"<td  class='tdbg'><input name='fdbw_wh' type='text'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ʣ�</strong></td> "
	 dwt.out"<td  class='tdbg'><input type='text' name='fdbw_jz' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;��</strong></td> "
	dwt.out"<td><select name='fdbw_lb' size='1'>"
	dwt.out"<option value='1'>һ</option>"
    dwt.out"<option value='2'>��</option>"
    dwt.out"</select></td></tr>"
	 
    dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>������ʽ��</strong></td>"
	dwt.out"<td><select name='fdbw_brxx' size='1'>"
	dwt.out"<option value='1'>��</option>"
    dwt.out"<option value='2'>��</option>"
    dwt.out"</select></td></tr>"
		dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='fdbw_tyqk' size='1'>"
	dwt.out"<option value='1'>Ͷ��</option>"
	dwt.out"<option value='2'>�߱�����</option>"
	dwt.out"<option value='3'>��ȱ��</option>"
    dwt.out"</select></td></tr>"

		dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��ʼʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_csdate' type='text' value="&now()&"  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf

dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>Ͷ��ʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_date' type='text' value="&now()&"  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf
 
 	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>ͣ��ʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_tydate' type='text' value="&now()&"  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf

	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    dwt.out"<td  class='tdbg'><input type='text' name='fdbw_bz'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
	  on error resume next
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from fdbw" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      'dim tydate
	 ' if request("fdbw_tydate")=" " then  tydate="0000-00-00"
	  
      rsadd("sscj")=Trim(Request("fdbw_sscj"))
      rsadd("wh")=request("fdbw_wh")
      rsadd("ssgh")=Trim(request("fdbw_gh"))
      rsadd("jz")=request("fdbw_jz")
      rsadd("lb")=request("fdbw_lb")
      rsadd("brxx")=request("fdbw_brxx")
	  rsadd("tyqk")=request("fdbw_tyqk")
      rsadd("bz")=request("fdbw_bz")
	  rsadd("date")=request("fdbw_date")
	  rsadd("csdate")=request("fdbw_csdate")
	  rsadd("tydate")=request("fdbw_tydate")
	  rsadd("update")=now()
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='fdbw.asp';</Script>"
end sub

sub saveedit()    
      on error resume next
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from fdbw where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connjg,1,3
      rsedit("sscj")=Trim(Request("fdbw_sscj"))
      rsedit("wh")=request("fdbw_wh")
      rsedit("ssgh")=Trim(request("fdbw_gh"))
      rsedit("jz")=request("fdbw_jz")
      rsedit("lb")=request("fdbw_lb")
      rsedit("brxx")=request("fdbw_brxx")
      rsedit("bz")=request("fdbw_bz")
	  rsedit("date")=request("fdbw_date")
	  rsedit("tyqk")=request("fdbw_tyqk")
	  rsedit("csdate")=request("fdbw_csdate")
	  rsedit("tydate")=request("fdbw_tydate")
      	  rsedit("update")=now()

	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from fdbw where id="&id
  rsdel.open sqldel,connjg,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub


sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from fdbw where id="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><br><form method='post' action='fdbw.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table  border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�༭�������±�</strong></div></td>    </tr>"
     dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     dwt.out"<td  class='tdbg'><input name='fdbw_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out"<input name='fdbw_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'>"
	 dwt.out"<strong>��&nbsp;&nbsp;�ţ�</strong></td>"

	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
     dwt.out"<td  class='tdbg'>"
	 	dwt.out"<select name='fdbw_gh' size='1' >"
     call formgh (rsedit("ssgh"),rsedit("sscj"))
    dwt.out"</select> ��û����Ҫ�Ĺ���װ��,����ϵ�����������Ӧװ�ù�������</td></tr>"

	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'>"
	 dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 dwt.out"<td  class='tdbg'><input name='fdbw_wh' type='text' value='"&rsedit("wh")&"'></td>    </tr>   "
	 
	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ʣ�</strong></td> "
	 dwt.out"<td  class='tdbg'><input type='text' name='fdbw_jz'  value='"&rsedit("jz")&"'></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;��</strong></td> "
	dwt.out"<td><select name='fdbw_lb' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("lb")=1 then dwt.out"selected"
	dwt.out">һ</option>"
    dwt.out"<option value='2'"
	if rsedit("lb")=2 then dwt.out"selected"
	dwt.out">��</option>"
    dwt.out"</select></td></tr>"
    dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='fdbw_tyqk' size='1'>"
	dwt.out"<option value='1'>Ͷ��</option>"
	dwt.out"<option value='2'>�߱�����</option>"
	dwt.out"<option value='3'>��ȱ��</option>"
    dwt.out"</select></td></tr>"
    dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>������ʽ��</strong></td>"
	dwt.out"<td><select name='fdbw_brxx' size='1'>"
	dwt.out"<option value='1'"
	if rsedit("brxx")=1 then dwt.out"selected"
	dwt.out">��</option>"
	dwt.out"<option value='2'"
	if rsedit("brxx")=2 then dwt.out"selected"
	dwt.out">��</option>"
    dwt.out"</select></td></tr>"
    
	 
   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��ʼʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_csdate' type='text' value='"&rsedit("csdate")&"' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf

   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>Ͷ��ʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_date' type='text' value='"&rsedit("date")&"' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf

   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>ͣ��ʱ�䣺</strong></td> "
   dwt.out"<td  class='tdbg'>"
   dwt.out"<input name='fdbw_tydate' type='text' value='"&rsedit("tydate")&"' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
   dwt.out"</td></tr>"& vbCrLf

	 
	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    dwt.out"<td  class='tdbg'><input type='text' name='fdbw_bz' value='"&rsedit("bz")&"'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
	       rsedit.close
       set rsedit=nothing
	

end sub


sub main()
    url="fdbw.asp"
    dim title
	sqlfdbw="SELECT * from fdbw "
	if keys<>"" then 
		sqlfdbw=sqlfdbw&" where wh like '%" &keys& "%' "
		title="-���� "&keys
		url="fdbw.asp?keyword="&keys
	end if 
	if sscjid<>"" then
		sqlfdbw=sqlfdbw&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
		url="fdbw.asp?sscj="&sscjid
	end if 
	if ssghid<>"" then
	    sqlfdbw=sqlfdbw&" where ssgh="&ssghid
	    title="-"&gh(ssghid)
		url="fdbw.asp?ssgh="&ssghid
	end if 
	
	if request("update")<>"" then 
		sqlfdbw=sqlfdbw&" ORDER BY update desc"
    else
		sqlfdbw=sqlfdbw&" ORDER BY SSCJ ASC,ssGH ASC,WH ASC"
	end if 
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:5px;'>��������</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	dim sql,rs,numb,totall,tyl,sqlty,sqlwjb,sqljb,sqlqx
	sql="select * from levelname where istq=false AND LEVELID<4"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		dwt.out "û����ӳ���"
	else
	   do while not rs.eof
		sql="SELECT count(id) FROM fdbw WHERE sscj="&rs("levelid")
		sqlty="SELECT count(id) FROM fdbw WHERE sscj="&rs("levelid")&" and tyqk=1"'Ͷ��
		sqlwjb="SELECT count(id) FROM fdbw WHERE sscj="&rs("levelid")&" and tyqk=3"'δ�߱�����
		sqljb="SELECT count(id) FROM fdbw WHERE sscj="&rs("levelid")&" and tyqk=2"'�߱�����

		sqlqx="SELECT count(id) FROM fdbw WHERE sscj="&rs("levelid")&" and tyqk=4"'ȡ��
		'message connjg.Execute(sqlty)(0)/connjg.Execute(sqlty)(0)
		numb=numb&sscjh_d(rs("levelid"))&":"&"<span style='color:#006600;'>"&connjg.Execute(sql)(0)&"/"&connjg.Execute(sqlty)(0)&"/"&connjg.Execute(sqljb)(0)&"</span>/<span style='color:#ff0000'>"&connjg.Execute(sqlwjb)(0)&"</span>/"&connjg.Execute(sqlqx)(0)&"&nbsp;&nbsp;"
	 
	rs.movenext
	loop
	end if 
	rs.close
	
	sql="SELECT count(id) FROM fdbw"
	totall= "<span style='color:#006600;'>"&connjg.Execute(sql)(0)&"</span>" 
	dwt.out "<div class='pre'><div align=left>����/Ͷ��/�߱�/δ�߱�/ȡ��  "&numb&"�ϼ�:"&totall&"</div></div>"& vbCrLf

	call search()
	
	set rsfdbw=server.createobject("adodb.recordset")
	rsfdbw.open sqlfdbw,connjg,1,1
	if rsfdbw.eof and rsfdbw.bof then 
	   message("δ�ҵ���ؼ�¼")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>" & vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>λ��</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>������ʽ</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>״̬</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��ʼʱ��</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>Ͷ��ʱ��</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>ͣ��ʱ��</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>��ע</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>ѡ��</div></td>" & vbCrLf
		dwt.out "    </tr>" & vbCrLf
	
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
			select case rsfdbw("lb")
			  case 1
				 lb="һ"
			  case 2 
				lb="��"
			end select	 
			select case rsfdbw("brxx")
			  case 1
				 brxx="��"
			  case 2 
				brxx="��"
			end select	 
			dim tyqk
			select case rsfdbw("tyqk")
			  case 1
				 tyqk="<span style='color:#006600'>Ͷ��</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>�߱�����</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>��ȱ��</span>"
			  case 4 
				tyqk="ȡ������"
			end select
if rsfdbw("tyqk")="" then typk=""
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
					dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>" & vbCrLf
					dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsfdbw("sscj"))&"</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&gh(rsfdbw("ssgh"))&"&nbsp;</td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"""
					if now()-rsfdbw("update")<7 then   dwt.out "bgcolor=""#FFFF00"""
					dwt.out ">"&rsfdbw("wh")&"&nbsp;</td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("jz")&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&lb&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&brxx&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=fdbw_whjl.asp?fdbwid="&rsfdbw("id")&">"&tyqk&"</a></div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("csdate")&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("date")&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("tydate")&"&nbsp;</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsfdbw("bz")&"&nbsp;</td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>" & vbCrLf
					call editdel(rsfdbw("id"),rsfdbw("sscj"),"fdbw.asp?action=edit&id=","fdbw.asp?action=del&id=")
					dwt.out "</div></td></tr>" & vbCrLf
					 RowCount=RowCount-1
			  rsfdbw.movenext
			  loop
			dwt.out "</table>" & vbCrLf
		    if sscjid<>"" or ssghid<>"" or keys<>"" then 
			  call showpage(page,url,total,record,PgSz)
			else
			  call showpage1(page,url,total,record,PgSz)
			end if  
		   end if
		   rsfdbw.close
		   set rsfdbw=nothing
			connjg.close
			set connjg=nothing
end sub





dwt.out "</body></html>"



sub search()
	dim sqlcj,rscj,sqlgh,rsgh
	dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='fdbw.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then  dwt.out "<a href='fdbw.asp?action=add'>��ӷ�������</a>"
	dwt.out "&nbsp;&nbsp;<a href='fdbw.asp?update=update'>�鿴����������</a>"
	dwt.out "  <input type='text' name='keyword'  size='20' maxlength='50' "
	if keys<>"" then 
	 dwt.out "value='"&keys&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='����������λ��'"
	 	dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if                 
	dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
	dwt.out "&nbsp;&nbsp;<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>��������ת����</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='fdbw.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid")  then dwt.out" selected"
			dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf

	dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>��װ����ת����</option>" & vbCrLf
	sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
		set rsgh=server.createobject("adodb.recordset")
		rsgh.open sqlgh,conn,1,1
		do while not rsgh.eof
			dwt.out"<option value='fdbw.asp?ssgh="&rsgh("ghid")&"'"
			if cint(request("ssgh"))=rsgh("ghid") then dwt.out" selected"
			dwt.out ">"&rsgh("gh_name")&"("&Connjg.Execute("SELECT count(id) FROM fdbw WHERE ssgh="&rsgh("ghid"))(0)&")</option>"& vbCrLf
		
			rsgh.movenext
		loop
		rsgh.close
		set rsgh=nothing
		dwt.out "     </select>	" & vbCrLf
		dwt.out "</form></div></div>" & vbCrLf

end sub





Call Closeconn
%>