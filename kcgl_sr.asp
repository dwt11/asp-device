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
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql
dim pagename

if Request("class")="" and request("sscj")="" then pagename="���̨��"
if Request("class")<>"" then 
 sql="SELECT * from kcclass where id="&Request("class")
 set rs=server.createobject("adodb.recordset")
 rs.open sql,connkc,1,1
 if rs.eof and rs.bof then 
    else
      pagename="���̨��--"&rs("name")
  end if 
 rs.close
 set rs=nothing
end if 

if request("sscj")<>"" then 
    sql="SELECT * from levelname where levelclass=1 and levelid="&Request("sscj")
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then 
    else
      pagename="���̨��--"&rs("levelname")
  end if 
	rs.close
	set rs=nothing
end if

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>���̨�˹���ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checksearch(){" & vbCrLf
response.write "  if(document.searchform.qsdate.value==''){" & vbCrLf
response.write "      alert('��ʼ���ڲ���Ϊ�գ�');" & vbCrLf
response.write "  document.searchform.qsdate.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
 
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>"&pagename&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'>"
dim sqlclass,rsclass '��ʾ  �����
sqlclass="SELECT * from class"
set rsclass=server.createobject("adodb.recordset")
rsclass.open sqlclass,connkc,1,1
do while not rsclass.eof
   response.write "<strong>"&rsclass("name")&":</strong>&nbsp;&nbsp;&nbsp;&nbsp;"
   sql="SELECT * from kcclass where class="&rsclass("id")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,connkc,1,1
   do while not rs.eof
	  response.write "<a href=kcgl_sr.asp?class="&rs("id")&">"&rs("name")&"</a>&nbsp;|&nbsp;"& vbCrLf
   rs.movenext
   loop
   rs.close
   set rs=nothing
   response.write "<br>"

rsclass.movenext
loop
rsclass.close
set rsclass=nothing   
response.write "</td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf
call search()

if Request("action")="delsrinfo" then call delsrinfo    'ɾ�������Ϣ




dim sqlbody,rsbody,xh


if request("class")="" and request("sscj")="" then 
   url="kcgl_sr.asp"
   sqlbody="SELECT * from sr order by id DESC"
end if 
if request("class")<>"" then 
   url="kcgl_sr.asp?class="&request("class")
   sqlbody="SELECT * from sr where class="&request("class")&" order by id DESC"
end if 

if request("sscj")<>"" then 
   url="kcgl_sr.asp?sscj="&request("sscj")
   sqlbody="SELECT * from sr where sscj="&request("sscj")&" order by id DESC"
end if 


  'sqlbody="SELECT * from sr order by id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">��������</p>" 
  else
     record=rsbody.recordcount
     if Trim(Request("PgSz"))="" then
        PgSz=20
     ELSE 
        PgSz=Trim(Request("PgSz"))
     end if 
     rsbody.PageSize = Cint(PgSz) 
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
     rsbody.absolutePage = page
     start=PgSz*Page-PgSz+1
     rowCount = rsbody.PageSize
  
     response.write "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     response.write "<tr class=""title"">"
     response.write "<td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>���</strong></div></td>"
     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��Դ</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ͺ�</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��λ</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���ʱ��</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ע</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ ��</strong></div></td>"
     response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        xh=xh+1
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("wpid")&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sscj"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&dclass(rsbody("class"))&"-"&kcclass(rsbody("class"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("lytxt")&"&nbsp;</div></td>"

        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&rsbody("name")&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&rsbody("xhgg")&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dw")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dmoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("numb")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("amoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("sr_year")&"-"&rsbody("sr_month")&"-"&rsbody("sr_day")&"</div></td>"
		response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("bz")&"&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
       if session("level")=rsbody("sscj") or session("level")=0 then 
	    response.write "<a href=kcgl_sr.asp?action=delsrinfo&id="&rsbody("id")&" onClick=""return confirm('ȷ��Ҫɾ��������¼��');"">ɾ��</a>"
     else
        response.write "&nbsp;"
     end if 
	   response.write "</div></td></tr>"
       dim totalamoney '�ϼ�ҳ����ܽ��
	   totalamoney=totalamoney+rsbody("amoney")
	    RowCount=RowCount-1
    rsbody.movenext
    loop
           response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color=#FF0000>�ϼ�</font></div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
       response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color=#FF0000>"&totalamoney&"</font>&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"

       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td></tr>"

   response.write "</table>"
  
  
  if request("class")="" and request("sscj")="" then 
     call showpage1(page,url,total,record,PgSz)
  else
     call showpage(page,url,total,record,PgSz)
  end if 
 end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing


response.write "</body></html>"
sub delsrinfo()
  dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from sr where id="&request("id")
  rsdel.open sqldel,connkc,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
 set rsdel=nothing  
end sub

'���ڿ���ӷ���������ʾ
Function kcclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from kcclass where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connkc,1,1
    do while not rsname.eof
	    kcclass=rsname("name")
		rsname.movenext
	loop
	rsname.close
	set rsname=nothing
end Function 

'������ʾ���������� 
Function dclass(classid)
	dim sqlname,rsname
	dim sqlz,rsz
	sqlz="SELECT * from kcclass where id="&classid
    set rsz=server.createobject("adodb.recordset")
    rsz.open sqlz,connkc,1,1
    'do while not rsz.eof
	 '   kcclass=rsname("name")
		'rsname.movenext
	'loop
	   sqlname="SELECT * from class where id="&rsz("class")
       set rsname=server.createobject("adodb.recordset")
       rsname.open sqlname,connkc,1,1
       'do while not rsname.eof
	    dclass=rsname("name")
		'rsname.movenext
	'loop
	rsname.close
	set rsname=nothing
	rsz.close
	set rsz=nothing
end Function 


'ѡ��༭������\ɾ����
sub editdel(id,sscj)
 if session("level")=sscj or session("level")=0 then 
    response.write "�༭&nbsp;"
	response.write "<a href=kcgl.asp?action=fc&id="&id&">����</a>&nbsp;"
	response.write "<a href=kcgl.asp?action=del&id="&id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a>"
 else
    response.write "&nbsp;"
 end if 
end sub


sub search()
dim sqlcj,rscj
response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
response.write "<tr class='tdbg'>" & vbCrLf

'����������
response.write "  <form method='Get' name='SearchForm' action='kcgl_search.asp'>" & vbCrLf
response.write "   <td>  <strong>�����Ϣ������</strong>" & vbCrLf
response.write "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"">" & vbCrLf
response.write "  <input type='submit' name='Submit'  value='����'>" & vbCrLf
response.write "  <input type='hidden' name='action' value='sr'>" & vbCrLf
response.write "</td></form>"

'��ʱ������
response.write "  <form method='Get' name='searchform' action='kcgl_search.asp'  onsubmit='javascript:return checksearch();'>" & vbCrLf
response.write "<td><strong>ʱ�������:</strong><script language=javascript src='/js/popcalendar.js'></script>"
response.write"<input name='qsdate' type='text' value="&now()&"  size=9>"
response.write"<a href='#' onClick=""popUpCalendar(this,qsdate, 'yyyy-mm-dd'); return false;"">"
response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a>"
response.write"����<input name='zzdate' type='text' value="&now()&" size=9>"
response.write"<a href='#' onClick=""popUpCalendar(this,zzdate, 'yyyy-mm-dd'); return false;"">"
response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a>"
response.write "  <input type='submit' name='submit'  value='����'>" & vbCrLf
response.write "  <input type='hidden' name='action' value='sr'>" & vbCrLf
response.write "</td></form>"

response.write "<td><font color='0066CC'> ������������ݣ�</font>"
response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "	       <option value=''>��������ת����</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='kcgl_sr.asp?sscj="&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	response.write "     </select>	" & vbCrLf
response.write "</td>  </tr></table>" & vbCrLf
end sub


Call CloseConn
%>