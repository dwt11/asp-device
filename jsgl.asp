<%@language=vbscript codepage=936 %>
<%
Option Explicit
response.buffer = True
Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<title>ϵͳ������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Style.css">
<style type="text/css">
<!--
.style2 {
	color: #FFFFFF;
	font-size: 18pt;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="49"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#5C94E7">
  <tr>
    <td height="34" ><div align="center" class="style2" >����������ҳ</div></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="481"><img src="Images/main_03.gif" width="481" height="35"></td>
    <td align="right" background="Images/main_04bg.gif"><img src="Images/main_04.gif" width="68" height="35"></td>
    <td width="20" background="Images/main_04bg.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="134"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50" >��������������¹��ܣ�<a href="ylb.asp">��������</a>&nbsp;&nbsp;<a href="lsda.asp">��������</a>&nbsp;&nbsp;<a href="fdbw.asp">��������</a>
����̨��&nbsp;&nbsp;����ͼֽ&nbsp;&nbsp;���޼�¼&nbsp;&nbsp;�豸��ת��&nbsp;&nbsp;��ѵ���� </td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;1���������� ȫ���豸��������Ϊ��<a href="ylb.asp?lxclassid=1">��ӵ�ѹ����</a> <a href="ylb.asp?lxclassid=2">������</a> <a href="ylb.asp?lxclassid=3">ת����</a> <a href="ylb.asp?lxclassid=4">���ڷ�����</a> <a href="ylb.asp?lxclassid=5">��ŷ�</a> <a href="ylb.asp?lxclassid=6">�͵ص�����</a> <a href="ylb.asp?lxclassid=7">ת��̽ͷ</a> <a href="ylb.asp?lxclassid=8">����һ��Ԫ��</a> <a href="ylb.asp?lxclassid=9">����һ��Ԫ��</a> <a href="ylb.asp?lxclassid=10">����̽ͷ</a> <a href="ylb.asp?lxclassid=11">����</a> <a href="ylb.asp?lxclassid=12">�յ�</a> <a href="ylb.asp?lxclassid=13">Ƥ����</a> <a href="ylb.asp?lxclassid=14">���ڷ�</a> <a href="ylb.asp?lxclassid=15">�綯ִ�л���</a>ʮ������ࡣ</td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;2���������� ȫ������ͳ�ƻ��ܡ�</td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;3���������� ȫ�������������±�ͳ�ƻ��ܡ�</td>
  </tr>
</table>

</body>
</html>
<%



Call CloseConn
%>