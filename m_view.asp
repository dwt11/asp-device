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
dim sqlmessage,rsmessage,sql,rs,m_username

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ�ڲ��ʼ�</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
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

sqlmessage="SELECT * from message where id="&request("id")
set rsmessage=server.createobject("adodb.recordset")
rsmessage.open sqlmessage,connd,1,1
response.write "<div align=center>�鿴�ʼ�����</div>"
response.write "<table width='70%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class=""tdbg"">"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>�����ˣ�&nbsp;&nbsp;</strong>"& vbCrLf

sql="SELECT * from userid where id="&rsmessage("formid")&" ORDER BY id DESC"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
m_username=rs("username")
rs.close
set rs=nothing
response.write m_username&"&nbsp;&nbsp;&nbsp;&nbsp;<strong>ʱ�䣺</strong>&nbsp;&nbsp;"&rsmessage("date")&"</td>"
response.write "  </tr>  "& vbCrLf

response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='50%' height='30'>"&rsmessage("body")&"</td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

rsmessage.close
set rsmessage=nothing
connd.close
set connd=nothing





response.write "</body></html>"

Call CloseConn
%>