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
Dwt.out "alert('��ѡ���������䣡');" & vbCrLf
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
	ssbz=request("ssbz")
	url="zjqk_post.asp?action=zjpost&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj&"&ssbz="&ssbz
	
	if zjmonth=0 then
	   zjmonth_d="����"
	else
	   zjmonth_d=zjmonth&"��"
	end if       
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&zjyear&"��-"&zjmonth_d&" "&sscjh(sscj)&" "&ssbzh(ssbz)&" �ܼ�̨��</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf
'    Dwt.Out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
'	Dwt.out "<input type='button' name='Submit'  onclick=""window.location.href='tocsv.asp?action=zjtz&sscj="&sscj&"&ssbz="&ssbz&"&zjyear="&zjyear&"&zjmonth="&zjmonth&"&titlename=�ܼ�̨��'"" value='������ҳ���ݵ�EXCEL'>"
'    dwt.out "</div></div>"	

	'if zjmonth<>0 then sql="SELECT * from zjtz where (year(sczjdate)="&zjyear&"  or year(sczjdate)="&zjyear&"-jdzq/12) and isdx=false and month(sczjdate)="&zjmonth&" and sscj="&sscj&" and ssbz="&ssbz&" ORDER BY id aSC "
	'if zjmonth=0 then sql="SELECT * from zjtz where (dxzjyear="&zjyear&"  or dxzjyear="&zjyear&"-jdzq/12) and isdx and sscj="&sscj&" and ssbz="&ssbz&" ORDER BY id aSC "


if zjmonth<>0 then sql="SELECT * from zjtz where (year(dateadd('m',jdzq,sczjdate))="&zjyear&" or year(sczjdate)="&zjyear&") and isdx=false and (month(dateadd('m',jdzq,sczjdate))="&zjmonth&"  or month(sczjdate)="&zjmonth&") and sscj="&sscj&" and ssbz="&ssbz&" and jdzq<>0 ORDER BY id aSC "
	if zjmonth=0 then sql="SELECT * from zjtz where (dxzjyear="&zjyear&"  or dxzjyear="&zjyear&"-jdzq/12) and isdx and sscj="&sscj&" and ssbz="&ssbz&" and jdzq<>0 ORDER BY id aSC "

	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
		message "δ�ҵ��������" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
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
					Dwt.Out "      <td  class='x-td'>"&uCase(rs("wh"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("jdzq")&"&nbsp;</Div></td>" & vbCrLf
	
					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					jdzq=rs("jdzq")/12
					
			'�ϴ��ܼ�����
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"				   
			if rs("isdx") then 
			      if year(rs("sczjdate"))=zjyear then Dwt.out rs("dxzjyear")&"-����"
			      if year(rs("sczjdate"))<>zjyear then Dwt.out rs("dxzjyear")+jdzq&"-"&"����"
			else
			      'if year(rs("sczjdate"))=zjyear then Dwt.out rs("sczjdate")
			     
				 ' if year(rs("sczjdate"))<>zjyear then Dwt.out year(rs("sczjdate"))+jdzq&"-"&month(rs("sczjdate"))&"-"&day(rs("sczjdate"))
			   Dwt.out  dateadd("m",rs("jdzq"),rs("sczjdate"))

'Dwt.out rs("sczjdate")&"sdf"&zjyear
			end if 	 	 
			Dwt.out "</Div></td>" & vbCrLf
			 'Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rsscdate("zjinfo")&"</Div></td>" & vbCrLf
			
			dim sqlinfo,rsinfo
			dim c_text
			'�´��ܼ�����
			Dwt.Out "<td  class='x-td'><Div align=""center"">"

			
			if zjmonth<>0 and not rs("isdx") then sqlinfo="SELECT * from zjinfo where not isdx and year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("id")
			if zjmonth=0 and rs("isdx") then sqlinfo="SELECT * from zjinfo where isdx and dxzjyear="&zjyear&"  and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "δ�ܼ�"
				'if  (year(now())>=zjyear AND month(now())>zjmonth) or (zjyear>=year(now()) AND zjmonth>month(now())) then 
					'c_text="�ѹ���"
				'else	
					c_text="<a href=zjqk_post.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjdate="&zjyear&"-"&zjmonth&">���</a>  "
				'end if 

			    c_text=c_text&"  <a href=zjqk_post.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&">���ļƻ�����</a>"
			else
			    IF RSINFO("ISDX") THEN DWT.OUT RSINFO("DXZJYEAR")&"-����"
				IF NOT RSINFO("ISDX") THEN DWT.OUT RSINFO("zjdate")
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
  dwt.out "<a href=tocsv.asp?action=zjtz&titlename="&zjyear&"��"&zjmonth&"��"&sscjh_D(sscj)&ssbzh(ssbz)&"�ܼ�̨��&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj&"&ssbz="&ssbz&">����</a>"
   Dwt.Out "</Div>"
   end if
   Dwt.Out "</Div>"		   
   rs.close
   set rs=nothing
End Sub

'���ڱ��汾���ܼ���ɺ�������ܼ���
sub complete()
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from zjtz where id="&request("id")&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
        message("δ֪����")
   else
	   Dwt.out"<br><br><br><form method='post' action='zjqk_post.asp' name='form2' onsubmit='javascript:return complete();'>"
	   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
	   Dwt.out"<Div align='center'><strong>�ܼ�����д</strong></Div></td>    </tr>"
	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
	   Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&sscjh(rszjtz("sscj"))&"' size=10>&nbsp;<input disabled='disabled'  type='text' value='"&ssbzh(rszjtz("ssbz"))&"' size=8></td></tr>"& vbCrLf
		
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
		Dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
		Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("wh")&"></td>    </tr>   "
		 
		 
		 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
		Dwt.out"<td><input disabled='disabled' type='text' value="&zjclass(rszjtz("class"))&"></td></tr>"
		 
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("ggxh")&"></td>    </tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("clfw")&"></td>    </tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("jdzq")&"></td></tr>"
'		if rszjtz("isdx") then 
'			zjyear=rszjtz("dxzjyear")+(rszjtz("jdzq")/12)		
'		else
'    		zjyear=year(rszjtz("sczjdate"))+(rszjtz("jdzq")/12)
'		end if 
		
'		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ܼ���ȣ�</strong></td>"
'		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&zjyear&"></td></tr>"
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
	Dwt.out "<input type='checkbox' name='isdx' "
	if rszjtz("isdx") then dwt.out "checked "
	dwt.out "onclick='zjtz_dxyear.disabled=!checked;zjtz_date.disabled=checked;'/>�Ƿ��Ǵ���"
	Dwt.out "<br/><select name='zjtz_dxyear'"
	if rszjtz("isdx")=false then dwt.out " disabled='disabled'"
	dwt.out ">" 
	for  i=year(now())-5 to year(now())+5
         Dwt.out "<option value="&i
		 if i=rszjtz("dxzjyear") then Dwt.out " selected"
	     Dwt.out ">"&i&"</option>"
	next
	Dwt.out "</select>�����ܼ����"
    Dwt.out"<br/><input name='zjtz_date' "
	if rszjtz("isdx") then dwt.out "disabled='disabled'"
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&request("zjdate")&"'/>�ճ��ܼ�����"		
		'Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td>"
'		 
'		 if zjmonth_d=0 then zjmonth_d="����"
'		 Dwt.out"<td width='80%' class='tdbg'><input disabled='disabled' type='text' value="&request("zjyear")&"-"&zjmonthname&"></td>    </tr>   "
'		'end if 
'		
'		Dwt.out"<input type='hidden' name=""zjyear"" value='"&request("zjyear")&"'>"
'		Dwt.out"<input type='hidden' name=""zjmonth"" value='"&request("zjmonth")&"'>"
'		Dwt.out"<input type='hidden' name=""sscj"" value='"&request("sscj")&"'>"
'		Dwt.out"<input type='hidden' name=""ssbz"" value='"&request("ssbz")&"'>"
'	
'		Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
'		 Dwt.out"<td width='80%' class='tdbg'>"
'		 Dwt.out"<select name=zjday>"
'		 dim i
'		 for i=1 to 31
'		  Dwt.out "<option value='"&i&"'"& vbCrLf
'		  if i=day(now()) then Dwt.out "selected"
'		  Dwt.out">"&i&"</option>"& vbCrLf
'		 next
'		 Dwt.out"</select></td></tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input name='zjinfo' type='text'></td>    </tr>   "
	
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
		Dwt.out"<td width='88%' class='tdbg'><input type='text' name='bz'></td></tr>  "   
	
		Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
		Dwt.out"<input name='action' type='hidden' id='action' value='completesave'> <input type='hidden' name='id' value='"&request("id")&"'>     <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'></td>  </tr>"
		Dwt.out"</table></form>"
		'Dwt.out request("sscj")&&
   end if 
end sub



sub completesave()
      dim rsadd,sqladd
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zjinfo" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("zjtzid")=Trim(Request("id"))
      if request("isdx")="on" then 
	     rsadd("dxzjyear")=request("zjtz_dxyear")
	     rsadd("isdx")=true
		 zjyear=request("zjtz_dxyear")
		 zjmonth=0
		' message request("isdx")&request("zjtz_dxyear")
	  else
	     rsadd("zjdate")=request("zjtz_date")
	     rsadd("isdx")=false
		 zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
	  end if 
      rsadd("bz")=request("bz")
      rsadd("zjinfo")=request("zjinfo")
	  rsadd.update
      rsadd.close
	 
	 
	  dim rsedit,sqledit
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zjtz where id="&ReplaceBadChar(Trim(request("id")))
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
      sscj=rsedit("sscj")
	  ssbz=rsedit("ssbz")
      rsedit.close
      set rsedit=nothing

	 
	  Dwt.out"<Script Language=Javascript>location.href='zjqk_post.asp?action=zjpost&sscj="&sscj&"&ssbz="&ssbz&"&zjyear="&zjyear&"&zjmonth="&zjmonth&"';</Script>"

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

	Dwt.out"<br/><form method='post' action='zjqk_post.asp' name='form1' onsubmit='javascript:return check();'>"
	Dwt.out "<table width='100%' >"& vbCrLf
	Dwt.out"<tr><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td> "
	Dwt.out"<td width='60%' class='tdbg'>"& vbCrLf
	Dwt.out "<select name='zjyear'>" & vbCrLf
	Dwt.out "<option value=''>ѡ�����</option>" & vbCrLf
	for i=year(now())-5 to year(now())+5
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=year(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
			'Dwt.out"<option value='"&i&"'>"&i&"</option>"& vbCrLf
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
		'if session("level")=rscj("levelid") then Dwt.out "selected"
		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
    Dwt.out "<select name='ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>ѡ��������</option>" & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""�ް���"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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
    Dwt.out "var temp=document.form1.ssbz" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//--></script>" & vbCrLf
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

Call Closeconn
%>