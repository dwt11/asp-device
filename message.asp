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
dim sqlmessage,rsmessage,title,record,pgsz,total,page,start,rowcount,xh,url,ii,m_body,m_username
dim rsadd,sqladd,messageid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,sql,rs,m_read
dim sqlusername,rsusername
url="message.asp"

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ�ڲ��ʼ�</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
%>
<script>

function setValue( src )
{
	if( document.form1.message_username.value == null ){
		return;	
	}

	var list = convert( src);
	if( document.form1.message_username.value != "")
	{
		if( document.form1.message_username.value.charAt( document.form1.message_username.value.lenth-1) == ",")
				list = compareResult( document.form1.message_username.value, list);
		else {
			list = compareResult( document.form1.message_username.value, list);
			if( list != "" )
				list =  "," + list;
		}
	}
	document.form1.message_username.focus();
	document.form1.message_username.value = document.form1.message_username.value + list;
}
function convert( email )
{
	var list = email;
	if( list.charAt( list.length-1) ==",")
		list = list.substring( 0,list.length -1);
	return list;
}
function compareResult(list1, list2)
{
	lt1 = list1.split(",");
	lt2 = list2.split(",");
	lt3=[];
	var index = 0;
	for(var i=0;i< lt2.length; i++)
	{
		var flag = false;
		for( var j=0;j<lt1.length; j++)
		{
			if( lt1[j] == lt2[i])
			{	flag = true;
				break;
			}
		}
		if( !flag )
			lt3[index] = lt2[i];
		index++;
	}
	return lt3.join( ",");
}

</script>

<%

response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>�ڲ��ʼ�ϵͳ</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='90' height='30'><strong>ϵͳ������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""message.asp?action=add"">д�ʼ�</a>&nbsp;|&nbsp;<a href=""message.asp?action=add"">������</a>&nbsp;|&nbsp;<a href=""message.asp"">������</a></td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

if Request("action")="add" then 
   call add
else
   if Request("action")="saveadd" then
      call saveadd
   else
		    if request("action")="del" then
			   call del
			else
			   call main 
			end if    
    end if  
end if 

sub add()
   response.write"<form method='post' action='message.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>д�ʼ�</strong></div></td>    </tr>"
	
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���ݣ�</strong></td>"      
    response.write"<td width='88%' class='tdbg'><textarea name='message_body' cols=""50"" rows=""10""></textarea></td></tr>  "   
	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ռ��ˣ�</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' size=""40"" name='message_username'>��������û���ѡ��������<br>"   
    sqlusername="SELECT * from userid  ORDER BY id DESC"
    set rsusername=server.createobject("adodb.recordset")
    rsusername.open sqlusername,conn,1,1
    if rsusername.eof and rsusername.bof then 
       response.write "<p align='center'>δ���յ��ʼ�</p>" 
    else
      do while not rsusername.eof 
            response.write "<a href='javascript:;' onclick='setValue("""&rsusername("username")&"<"&rsusername("id")&">"")'>"&rsusername("username")&"</a>&nbsp;&nbsp;"
        rsusername.movenext
      loop
    end if  
    rsusername.close
	
	response.write"</td></tr>"
	
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' ���� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from message" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("message_sscj"))
      rsadd("wh")=request("message_wh")
      rsadd("yt")=Trim(request("message_yt"))
      rsadd("ycjname")=request("message_ycjname")
      rsadd("cldw")=request("message_cldw")
      rsadd("clfw")=request("message_clfw")
      rsadd("lsl")=request("message_lsl")
      rsadd("lsh")=request("message_lsh")
      rsadd("tyzk")=request("message_tyzk")
      rsadd("zxzz")=request("message_zxzz")
      rsadd("bz")=request("message_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>window.alert('������������ɹ�');history.go(-2)</Script>"
end sub


sub saveedit()    
	  '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from message where messageid="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conn,1,3
      rsedit("sscj")=Trim(Request("message_sscj"))
      rsedit("wh")=request("message_wh")
      rsedit("yt")=Trim(request("message_yt"))
      rsedit("ycjname")=request("message_ycjname")
      rsedit("cldw")=request("message_cldw")
      rsedit("clfw")=request("message_clfw")
      rsedit("lsl")=request("message_lsl")
      rsedit("lsh")=request("message_lsh")
      rsedit("tyzk")=request("message_tyzk")
      rsedit("zxzz")=request("message_zxzz")
      rsedit("bz")=request("message_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>window.alert('�༭���������ɹ�');history.go(-2)</Script>"
end sub

sub del()
  messageid=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from message where id="&messageid
  rsdel.open sqldel,connd,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from message where id="&id
   rsedit.open sqledit,connd,1,1
   response.write"<form method='post' action='message.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>�༭��������</strong></div></td>    </tr>"
     
	 select case rsedit("sscj")
          case 1
             sscj="άһ"
          case 2 
        	sscj="ά��"
          case 3 
        	sscj="ά��" 
        end select	
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
     response.write"<td width='88%' class='tdbg'><input name='message_sscj'  disabled='disabled'  type='text' value='"&sscj&"'></td></tr>"& vbCrLf
     response.write"<input name='message_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='message_wh' type='text' value='"&rsedit("wh")&"'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='message_yt'  value='"&rsedit("yt")&"'></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>һ��Ԫ�����ƣ�</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='message_ycjname' value='"&rsedit("ycjname")&"'></td></tr> "
	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_cldw' value='"&rsedit("cldw")&"'></td></tr>  "   
   
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_clfw' value='"&rsedit("clfw")&"'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵL��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_lsl' value='"&rsedit("lsl")&"'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ֵH��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_lsh' value='"&rsedit("lsh")&"'></td></tr>  "   
   
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>Ͷ��״����</strong></td>"
	response.write"<td><select name='message_tyzk' size='1'>"
	response.write"<option value='1'>Ͷ��</option>"
    response.write"<option value='0'>��·</option>"
    response.write"</select></td></tr>"
	
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ִ��װ�ã�</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_zxzz' value='"&rsedit("zxzz")&"'></td></tr>  "   
	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='message_bz' value='"&rsedit("bz")&"'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"

end sub


sub main()
sqlmessage="SELECT * from message where userid="&session("id")&" ORDER BY id DESC"
set rsmessage=server.createobject("adodb.recordset")
rsmessage.open sqlmessage,connd,1,1
if rsmessage.eof and rsmessage.bof then 
response.write "<p align='center'>δ���յ��ʼ�</p>" 
else
response.write "<div align=""center"">�ռ���</div>"
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������</strong></div></td>"
response.write "      <td width=""65%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"
response.write "    </tr>"
           record=rsmessage.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsmessage.PageSize = Cint(PgSz) 
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
           rsmessage.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsmessage.PageSize
           do while not rsmessage.eof and rowcount>0
		xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&xh&"</div></td>"
                
				
				sql="SELECT * from userid where id="&rsmessage("formid")&" ORDER BY id DESC"
                set rs=server.createobject("adodb.recordset")
                rs.open sql,conn,1,1
				m_username=rs("username")
				 rs.close
                set rs=nothing
				response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&m_username&"</div></td>"
                
				m_body=rsmessage("body")
				if len(m_body)>38 then m_body=left(m_body,37)&"...."
				select case rsmessage("isread")
				   case 1
				response.write "      <td width=""60%"" style=""border-bottom-style: solid;border-width:1px""><a href=m_view.asp?id="&rsmessage("id")&">"&m_body&"</a>&nbsp;</td>"
				   case 0
				response.write "      <td width=""60%"" style=""border-bottom-style: solid;border-width:1px""><a href=m_view.asp?id="&rsmessage("id")&"><b>"&m_body&"</b></a>&nbsp;</td>"
				end select   
			    response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsmessage("date")&"</div></td>"
				response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=message.asp?action=del&id="&rsmessage("id")&" onClick=""return confirm('ȷ��Ҫɾ�����ʼ���');"">ɾ��</a></div></td></tr>"
				
                 RowCount=RowCount-1
          rsmessage.movenext
          loop
        response.write "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rsmessage.close
       set rsmessage=nothing
        connd.close
        set connd=nothing
end sub





response.write "</body></html>"

Call CloseConn
%>