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
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>�����������ҳ</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.Out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function check(){" & vbCrLf

Dwt.out "if(document.form1.sscj.value==''){" & vbCrLf
Dwt.out "alert('��ѡ��������λ��');" & vbCrLf
Dwt.out "document.form1.sscj.focus();" & vbCrLf
Dwt.out "return false;" & vbCrLf
Dwt.out "}" & vbCrLf

Dwt.out "}" & vbCrLf

Dwt.out "function complete(){" & vbCrLf

Dwt.out "if(document.form2.zjinfo.value==''){" & vbCrLf
Dwt.out "alert('�ܼ���δ��д��');" & vbCrLf
Dwt.out "document.form2.zjinfo.focus();" & vbCrLf
Dwt.out "return false;" & vbCrLf
Dwt.out "}" & vbCrLf

Dwt.out "}" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.Out"<script language=javascript src='/js/popselectdate.js'></script>"

Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dim sqlcj,rscj,i,ii,sqlbz,rsbz,sql,rs
    dim url,record,pgsz,total,page,start,rowcount
	dim zjyear,zjmonth
	dim sscj,ssbz
	dim zjmonth_d
action=request("action")

select case action 
   case "zjpost"
     call zjpost
   case ""
     call main
   case "complete"
     call complete
   case "completesave"
     call completesave
end select	  	 


Sub zjpost()
	dim zjmonth
	zjyear=cint(request("zjyear"))
	zjmonth=cint(request("zjmonth"))
    sscj=request("sscj")
	url="jycl_zj.asp?action=zjpost&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj
	zjmonth_d=zjmonth&"��"
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&zjyear&"��-"&zjmonth_d&" "&sscjh(sscj)&" ������������豸�ܼ�̨��</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf

	sql="SELECT * from jycltz where (year(sczjdate)="&zjyear&"  or year(sczjdate)="&zjyear&"-jdzq/12)  and month(sczjdate)="&zjmonth&" and sscj="&sscj&" ORDER BY id aSC "
	
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
		message "δ�ҵ��������" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>��λ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ͺ�</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ƻ���������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ʵ�ʼ�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
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
			Dwt.Out "     <td  Class='x-td'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
					Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))&"</Div></td>" & vbCrLf
					ssbz=rs("ssbz")
					Dwt.Out "      <td  class='x-td'>"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("jdzq")&"&nbsp;</Div></td>" & vbCrLf
	
					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					jdzq=rs("jdzq")/12
					
			'�ϴ��ܼ�����
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"				   
			
			      if year(rs("sczjdate"))=zjyear then Dwt.out rs("sczjdate")
			     
				  if year(rs("sczjdate"))<>zjyear then Dwt.out year(rs("sczjdate"))+jdzq&"-"&month(rs("sczjdate"))
			 
			Dwt.out "</Div></td>" & vbCrLf
			dim sqlinfo,rsinfo
			dim c_text
			'�´��ܼ�����
			Dwt.Out "<td  class='x-td'><Div align=""center"">"
		    sqlinfo="SELECT * from jycl_info where  year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "δ�ܼ�"
				
					c_text="<a href=jycl_zj.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&"&zjdate="&zjyear&"-"&zjmonth&">���</a>  "
				

			    c_text=c_text&"  <a href=jycl_zj.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&">���ļƻ�����</a>"
			else
			    
				dwt.out rsinfo("zjdate")
				dim jdjg
				if rsinfo("zjinfo")="" then
				   jdjg="δ��д�������"
				else
				   jdjg=rsinfo("zjinfo")
				end if       
				c_text="�ܼ���� "&jdjg
			end if 
			
			Dwt.out "</Div></td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'>"&rs("bz")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
			dwt.out c_text
			Dwt.Out "</Div></td></tr>" & vbCrLf
			c_text=""
			 RowCount=RowCount-1
	  rs.movenext
	  loop
	Dwt.Out "</table>" & vbCrLf
	   call showpage(page,url,total,record,PgSz)
   Dwt.Out "</Div>"
   end if
   Dwt.Out "</Div>"		   
   rs.close
   set rs=nothing
End Sub

'���ڱ��汾���ܼ���ɺ�������ܼ���
sub complete()
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from jycltz where id="&request("id")&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
        message("δ֪����")
   else
	   Dwt.out"<br><br><br><form method='post' action='jycl_zj.asp' name='form2' onsubmit='javascript:return complete();'>"
	   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
	   Dwt.out"<Div align='center'><strong>�ܼ�����д</strong></Div></td>    </tr>"
	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ�� </strong></td>"      
	   Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&sscjh(rszjtz("sscj"))&"' size=10>&nbsp;<input disabled='disabled'  type='text' value='"&ssbzh(rszjtz("ssbz"))&"' size=8></td></tr>"& vbCrLf		 
		 
		 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
		Dwt.out"<td><input disabled='disabled' type='text' value="&zjclass(rszjtz("class"))&"></td></tr>"
		 
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("ggxh")&"></td>    </tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("clfw")&"></td>    </tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("jdzq")&"></td></tr>"
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
    Dwt.out"<input name='zjtz_date' "
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&request("zjdate")&"'/>�ճ��ܼ�����"		
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input name='zjinfo' type='text'></td>    </tr>   "
	
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
		Dwt.out"<td width='88%' class='tdbg'><input type='text' name='bz'></td></tr>  "   
	
		Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
		Dwt.out"<input name='action' type='hidden' id='action' value='completesave'> <input type='hidden' name='id' value='"&request("id")&"'>     <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'></td>  </tr>"
		Dwt.out"</table></form>"
   end if 
end sub



sub completesave()
      dim rsadd,sqladd
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from jycl_info" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("zjtzid")=Trim(Request("id"))
	     rsadd("zjdate")=request("zjtz_date")
		 zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
	       rsadd("bz")=request("bz")
      rsadd("zjinfo")=request("zjinfo")
	  rsadd.update
      rsadd.close
	 
	 
	  dim rsedit,sqledit
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jycltz where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
	     rsedit("sczjdate")=request("zjtz_date")
	  
	  rsedit.update
      sscj=rsedit("sscj")
	  ssbz=rsedit("ssbz")
      rsedit.close
      set rsedit=nothing

	 
	  Dwt.out"<Script Language=Javascript>location.href='jycl_zj.asp?action=zjpost&sscj="&sscj&"&ssbz="&ssbz&"&zjyear="&zjyear&"&zjmonth="&zjmonth&"';</Script>"

end sub

sub main()
	Dwt.out "<br/><br/><br/><br/><br/>"
	dwt.out "<Div align='center'><Div class='x-dlg x-dlg-closable x-dlg-draggable x-dlg-modal' style=' WIDTH: 400px; HEIGHT: 198px'>"
	Dwt.out "  <Div class='x-dlg-hd-left'>"
	Dwt.out "    <Div class='x-dlg-hd-right'>"
	Dwt.out "      <Div class='x-dlg-hd x-unselectable'>�ܼ��豸��ѯ</Div>"
	Dwt.out "    </Div>"
	Dwt.out "  </Div>"
	Dwt.out "  <Div class='x-dlg-dlg-body' style='WIDTH: 400px;'><Div align=left>"

	Dwt.out"<br/><form method='post' action='jycl_zj.asp' name='form1' onsubmit='javascript:return check();'>"
	Dwt.out "<table width='100%' >"& vbCrLf
	Dwt.out"<tr><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td> "
	Dwt.out"<td width='60%' class='tdbg'>"& vbCrLf
	Dwt.out "<select name='zjyear'>" & vbCrLf
	Dwt.out "<option value=''>ѡ�����</option>" & vbCrLf
	for i=year(now())-5 to year(now())+5
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=year(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>��	" & vbCrLf
	Dwt.out "<select name='zjmonth'>" & vbCrLf
	Dwt.out "<option value=''>ѡ���·�</option>" & vbCrLf
	dwt.out "<option value=0>����</option>"
	for i=1 to 12
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=month(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>	" & vbCrLf
	Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	Dwt.out"<strong>�������䣺</strong></td>"
	Dwt.out "<td>" & vbCrLf
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	Dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'"& vbCrLf
		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
	Dwt.out "</td></tr>" & vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='zjpost'><input  type='submit' name='Submit' value='��ѯ' style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

	Dwt.out "  </Div></Div>"
	Dwt.out "</Div></Div>"
end sub



'���ڷ���������ʾ
Function zjclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from jycl_class where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    zjclass=rsname("name")
	end if 
	rsname.close
	set rsname=nothing
end Function

Call Closeconn
%>