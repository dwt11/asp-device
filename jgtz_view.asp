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
<!--#include file="inc/imgcode.asp"-->

<%
dim sqljgtz,rsjgtz,sql,rs,m_username,sscj,sqljgtz_bj,rsjgtz_bj

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ����̨�˹���ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>�鿴����̨������</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""jgtz.asp"">����̨����ҳ</a>&nbsp;|&nbsp;<a href=""jgtz.asp?action=add"">��Ӽ�����Ŀ</a>"& vbCrLf
sqljgtz="SELECT * from jgtz where id="&request("id")
set rsjgtz=server.createobject("adodb.recordset")
rsjgtz.open sqljgtz,connjg,1,1
if session("level")=rsjgtz("sscj") or session("level")=0 then 
    response.write "|&nbsp;<a href=""jgtz_bj.asp?action=add&jgtzid="&request("id")&""">��ӱ���</a>"
 else
    response.write "&nbsp;"
 end if 
rsjgtz.close
set rsjgtz=nothing

response.write "  </td></tr>"& vbCrLf

response.write "</table>"& vbCrLf

sqljgtz="SELECT * from jgtz where id="&request("id")
set rsjgtz=server.createobject("adodb.recordset")
rsjgtz.open sqljgtz,connjg,1,1
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='100%' height='10'  class=jgtz_tdbg colspan=8><div align=center>"&rsjgtz("name")&"</div></td>  </tr>"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='6%'>����ˣ�</td><td width='10%'>"&rsjgtz("tcr")&"&nbsp;</td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'>���ʱ�䣺</td><td width='10%'>"&rsjgtz("tcdate")&"&nbsp;</td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='10%'>��ĿͶ�ʣ�</td><td width='10%'>"&rsjgtz("xmtz")&"&nbsp;</td>"& vbCrLf
response.write"<td class=jgtz_tdbg width='10%'>��Լ���ʽ�</td><td width='6%'>"&rsjgtz("jyjjz")&"&nbsp;</td></tr>" &vbcrlf
response.write "</table>"& vbCrLf

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='10%' ><div align=center>��<br>��<br>ԭ<br>��</div></td><td width='95%' >"&rsjgtz("jgyy")&"&nbsp;</td>"& vbCrLf
response.write "</table>"& vbCrLf

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='10%' ><div align=center>��<br>��<br>��<br>��</div></td><td width='95%' >"&imgCode(rsjgtz("jgfa"))&"&nbsp;</td>"& vbCrLf
response.write "</table>"& vbCrLf

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='9%'><div align=center>����ʱ�䣺</div></td><td width='10%'><div align=center>"&rsjgtz("pf_date")&"&nbsp;</div></td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'><div align=center>���������</div></td><td width='10%'><div align=center>"&rsjgtz("pf_qk")&"&nbsp;</div></td>"& vbCrLf
response.write "<tr>"& vbCrLf
response.write "</table>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='9%'><div align=center>ʵʩ���䣺</div></td><td width='10%'><div align=center>"&sscjh(rsjgtz("sscj"))&"&nbsp;</div></td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'><div align=center>ʵʩʱ�䣺</div></td><td width='10%'><div align=center>"&rsjgtz("ssdate")&"&nbsp;</div></td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'><div align=center>ʵʩ�����ˣ�</div></td><td width='10%'><div align=center>"&rsjgtz("ssname")&"&nbsp;</div></td>"& vbCrLf
response.write "<tr>"& vbCrLf
response.write "</table>"& vbCrLf

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='10%' ><div align=center>ʵ<br>ʩ<br>��<br>��</div></td><td width='95%' >"&rsjgtz("ssqk")&"&nbsp;</td>"& vbCrLf
response.write "</table>"& vbCrLf

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='10%' ><div align=center>��<br>��<br>Ч<br>��</div></td><td width='95%' >"&rsjgtz("jgxg")&"&nbsp;</td>"& vbCrLf
response.write "</table>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='9%'><div align=center>���ʱ�䣺</div></td><td width='10%'><div align=center>"&rsjgtz("wc_date")&"&nbsp;</div></td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'><div align=center>��������</div></td><td width='10%'><div align=center>"&rsjgtz("wc_qk")&"&nbsp;</div></td>"& vbCrLf
response.write "<tr>"& vbCrLf
response.write "</table>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' >"& vbCrLf
response.write "<tr class=tdbg><td class=jgtz_tdbg width='9%'><div align=center>����ʱ�䣺</div></td><td width='10%'><div align=center>"&rsjgtz("jd_date")&"&nbsp;</div></td>"& vbCrLf
response.write "<td class=jgtz_tdbg width='9%'><div align=center>�����ȼ���</div></td><td width='10%'><div align=center>"&rsjgtz("jd_dj")&"&nbsp;</div></td>"& vbCrLf
response.write "<tr>"& vbCrLf
response.write "</table>"& vbCrLf

sqljgtz_bj="SELECT * from jgtz_bj where jgtzid="&request("id")
set rsjgtz_bj=server.createobject("adodb.recordset")
rsjgtz_bj.open sqljgtz_bj,connjg,1,1
if rsjgtz_bj.eof and rsjgtz_bj.bof then 
response.write "<p align='center'>δ��ӱ���</p>" 
else
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr>" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%"" class=jgtz_tdbg><div align=""center"">���</div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%"" class=jgtz_tdbg><div align=""center"">��������</div></td>"
response.write "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px"" class=jgtz_tdbg><div align=""center"">�����ͺ�</div></td>"
response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px"" class=jgtz_tdbg><div align=""center"">����</div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px"" class=jgtz_tdbg><div align=""center"">��ע</div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px"" class=jgtz_tdbg><div align=""center"">ѡ��</div></td>"
response.write "    </tr>"
            do while not rsjgtz_bj.eof 
     response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center"">"&rsjgtz_bj("id")&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center"">"&rsjgtz_bj("bj_name")&"</div></td>"
                response.write "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px"">"&rsjgtz_bj("bj_xh")&"&nbsp;</div></td>"
                response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz_bj("bj_sl")&"&nbsp;</div></td>"
                response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz_bj("bj_bz")&"&nbsp;</div></td>"
                response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				call editdel(rsjgtz_bj("id"),rsjgtz("sscj"),"jgtz_bj.asp?action=edit&jgtzid="&request("id")&"&id=","jgtz_bj.asp?action=del&id=")
				
                response.write "</div></td></tr>"
         rsjgtz_bj.movenext
          loop
        response.write "</table>"

end if 

rsjgtz_bj.close
set rsjgtz_bj=nothing
rsjgtz.close
set rsjgtz=nothing



response.write "</body></html>"
Call Closeconn
%>