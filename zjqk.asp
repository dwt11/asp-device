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
<%response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title> �����������ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dim sqlcj,rscj,i,ii,sqlbz,rsbz,sql,rs
if Request("action")="zjinfo" then call zjinfo
if request("action")="complete" then call complete
if request("action")="completesave" then call completesave


sub zjinfo()
'************************************�㷨

'��ZJTZ���б����������ܼ�ı�
'����춨������0��1��ͣ��OR���ܼ죩����
'��ZJINFO���д�����ID��ʼ���ұ���������ZJTZID�ı�
'���ZJYEAR���ύ�����ģ�������ƻ� �·ݡ����ڡ����ΪZJINFO���д������
'���ZJYEAR�����ύ�����ģ�������ƻ� �·�Ϊ�ύ���������ݣ����Ҽ������ںͽ��Ϊ�հ�

'�˶δ����д���
'***********************************88
dim zjinfoor    '�����ж��Ƿ��ҵ���Ӧ���ܼ���Ϣ
	zjinfoor=0
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from zjtz where sscj="&cint(request("sscj"))&" and ssbz="&cint(request("ssbz"))&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
      dim text
	 zjmonth=request("zjmonth")
	 if cint(zjmonth)=0 then 
	  zjmonthname="����"
	  text="δ�ҵ� "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonthname&"   �ܼ����"
	 else
	  text="δ�ҵ� "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonth&"��   �ܼ����"
	 end if  
	  call message(text)
   else
      response.write "<table height=50 width=""100%"" border=""0"" align=""center"" cellpadding=""0""><tr><td height=40><font size=""5""><div align=center>"
	  if cint(request("zjmonth"))=0 then
	     zjmonth="����"
	     response.write sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonthname&"   �ܼ����"
	  else
   	     zjmonth=cint(request("zjmonth"))
		 response.write sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonth&"��  �ܼ����"
	  end if    
	  response.write "</div></font></td></tr></table>"
	  response.write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
      response.write "<tr class=""title"">"  & vbCrLf
      response.write "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>���</strong></div></td>" & vbCrLf
      response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>λ��</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ͺ�</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������Χ</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�ƻ������·�</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�������</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ע</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>&nbsp;</strong></div></td>" & vbCrLf
      response.write "    </tr>" & vbCrLf
      do while not rszjtz.eof
          dim jdzq  '�춨�����ж�
		  dim jdyear '�춨���ڻ���Ϊ��
		  jdzq=rszjtz("jdzq")
		  if jdzq=0 then 
			  'response.write "<td><font color=#ff0000><div align=center>ͣ��</div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
		  else
		      if jdzq=1 then 
    		      'response.write "<td><font color=#ff0000><div align=center>���ܼ�</div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
			  else
				  jdyear=jdzq/12
		          sqlscdate="SELECT * from zjinfo where zjtzid="&rszjtz("id")&" ORDER BY id DESC"
				  'zjyear="&request("zjyear")-jdyear&" and zjmonth="&request("zjmonth")
                  set rsscdate=server.createobject("adodb.recordset")
                  rsscdate.open sqlscdate,connzj,1,1
                  if rsscdate.eof and rsscdate.bof then 
                       response.write "<td><div align=center>δ�ҵ�����,�����ܼ�̨������Ӵ˱�ĳ����ܼ�����</div></td></tr>" 
                  else
					   if rsscdate("zjyear")=cint(request("zjyear")) and rsscdate("zjmonth")=cint(request("zjmonth"))  then
                              response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                              response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rszjtz("id")&"</div></td>" & vbCrLf
                              response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_D(rszjtz("sscj"))&ssbzh(rszjtz("ssbz"))&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&zjclass(rszjtz("class"))&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("wh")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("ggxh")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("clfw")&"&nbsp;</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("jdzq")&"&nbsp;</div></td>" & vbCrLf
				              if rsscdate("zjmonth")=0 then 
							     zjmonthname="����"
							  else
							     zjmonthname=rsscdate("zjmonth")
							  end if  
							  response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjyear")&"-"&zjmonthname&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjday")&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjinfo")&"</div></td>" & vbCrLf
		                      response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("bz")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>&nbsp;</div>" & vbCrLf
                              response.write "</td></tr>" & vbCrLf
						zjinfoor=1
						else 
							  if rsscdate("zjyear")=cint(request("zjyear"))-jdyear and rsscdate("zjmonth")=cint(request("zjmonth"))  then
                                     response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                                     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rszjtz("id")&"</div></td>" & vbCrLf
                                     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_D(rszjtz("sscj"))&ssbzh(rszjtz("ssbz"))&"</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&zjclass(rszjtz("class"))&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("wh")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("ggxh")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("clfw")&"&nbsp;</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("jdzq")&"&nbsp;</div></td>" & vbCrLf
				                     if request("zjmonth")=0 then 
							            zjmonthname="����"
							         else
							            zjmonthname=request("zjmonth")
							         end if  
				                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&request("zjyear")&"-"&zjmonthname&"</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
		                             response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("bz")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center><a href=zjqk.asp?action=complete&id="&rszjtz("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&">���</aS></div>" & vbCrLf
                                     response.write "</td></tr>" & vbCrLf
                              'else
							         'response.write "<td><div align=center>δ�ҵ��������</div></td></tr>" 
							  						zjinfoor=1
							  end if 
						end if 	  
				end if 
			    rsscdate.close
		     end if 
	     end if   
    rszjtz.movenext
 
 loop
    response.write "</table>" & vbCrLf
 
 '�ж������ѭ���Ƿ��ҵ�������ݣ���������Ϣ��ʾ������ҵ������Ϣ������������ͷ���
  if zjinfoor=0 then 
   if cint(request("zjmonth"))=0 then 
	  zjmonthname="����"
	  text="δ�ҵ� "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonthname&"   �ܼ����"
	 else
	  text="δ�ҵ� "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"��"&zjmonth&"��   �ܼ����"
	 end if  
	  call message(text)
   else
   			response.write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  class='border'><tr class='tdbg'><td><div align=right>"
			response.write "<input type='button' name='Submit'  onclick=""window.location.href='tocsv.asp?action=zjtz&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&"&titlename=�ܼ�̨��'"" value='�����������ݵ�EXCEL'>"
			
			response.write "</div></td></tr></table>"

   
   end if 	  
 
 
   end if
   rszjtz.close
   set rszjtz=nothing
end sub
response.write "</body></html>"


'���ڱ��汾���ܼ���ɺ�������ܼ���
sub complete()
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from zjtz where id="&request("id")&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
        message("δ֪����")
   else
   response.write"<br><br><br><form method='post' action='zjqk.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�ܼ�����д</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
   response.write"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&sscjh(rszjtz("sscj"))&"' size=10>&nbsp;<input disabled='disabled'  type='text' value='"&ssbzh(rszjtz("ssbz"))&"' size=8></td></tr>"& vbCrLf
	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("wh")&"></td>    </tr>   "
	 
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
	response.write"<td><input disabled='disabled' type='text' value="&zjclass(rszjtz("class"))&"></td></tr>"
	 
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("ggxh")&"></td>    </tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("clfw")&"></td>    </tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("jdzq")&"></td></tr>"
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td>"
    'dim jdyear,zjyear
	'jdyear=rszjtz("jdzq")/12
    'sqlscdate="SELECT * from zjinfo where zjtzid="&rszjtz("id")&" ORDER BY id DESC"
    'set rsscdate=server.createobject("adodb.recordset")
    'rsscdate.open sqlscdate,connzj,1,1
    'if rsscdate.eof and rsscdate.bof then 
     '   response.write "<td><div align=center>δ�ҵ�����,�����ܼ�̨������Ӵ˱�ĳ����ܼ�����</div></td></tr>" 
     'else
	 'zjyear=rsscdate("zjyear")+jdyear
	 zjmonthname=request("zjmonth")
	 if zjmonthname=0 then zjmonthname="����"
	 response.write"<td width='80%' class='tdbg'><input disabled='disabled' type='text' value="&request("zjyear")&"-"&zjmonthname&"></td>    </tr>   "
    'end if 
	
	response.write"<input type='hidden' name=""zjyear"" value='"&request("zjyear")&"'>"
	response.write"<input type='hidden' name=""zjmonth"" value='"&request("zjmonth")&"'>"
	response.write"<input type='hidden' name=""sscj"" value='"&request("sscj")&"'>"
	response.write"<input type='hidden' name=""ssbz"" value='"&request("ssbz")&"'>"

	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	 response.write"<td width='80%' class='tdbg'>"
	 response.write"<select name=zjday>"
	 dim i
	 for i=1 to 31
	  response.write "<option value='"&i&"'"& vbCrLf
	  if i=day(now()) then response.write "selected"
	  response.write">"&i&"</option>"& vbCrLf
	 next
	 response.write"</select></td></tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='zjinfo' type='text'></td>    </tr>   "

	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='completesave'> <input type='hidden' name='id' value='"&request("id")&"'>     <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back(-1)"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
    'response.write request("sscj")&&
   end if 
end sub



sub completesave()
      dim rsadd,sqladd
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zjinfo" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("zjtzid")=Trim(Request("id"))
      rsadd("zjyear")=cint(Request("zjyear"))
	  rsadd("zjmonth")=cint(request("zjmonth"))
      rsadd("zjday")=request("zjday")
      rsadd("bz")=request("bz")
      rsadd("zjinfo")=request("zjinfo")
	  rsadd.update
rsadd.close
	  response.write"<Script Language=Javascript>location.href='zjqk.asp?action=zjinfo&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&"';</Script>"

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