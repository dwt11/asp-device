<%@language=vbscript codepage=936 %>
<%
Option Explicit%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sqlghname,rsghname,ghname
dim sqlbody,rsbody
sqlghname="SELECT  * from ghclass where ghid="&request("ghid")
    set rsghname=server.createobject("adodb.recordset")
    rsghname.open sqlghname,conn,1,1
    if rsghname.eof and rsghname.bof then 
       response.write "<p align='center'>暂无内容</p>" 
    else
       ghname=rsghname("ghname")
	 end if 
	 rsghname.close
	 set rsghname =nothing
	 
%>

<html>
<head>
<title><%=ghname%>设备技术档案列表页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td class="topbg"><div align="center"><%=ghname%>设备技术档案列表页</div></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="481"><img src="Images/main_03.gif" width="481" height="35"></td>
    <td align="right" background="Images/main_04bg.gif"><img src="Images/main_04.gif" width="68" height="35"></td>
    <td width="20" background="Images/main_04bg.gif">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="426"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="106">&nbsp;</td>
    <td width="839" valign="top">　　
    <%
	 sqlbody="SELECT * from body where ghid="&request("ghid")
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>暂无内容</p>" 
    else
       do while not rsbody.eof 
           Response.Write "<a href=body.asp?bodyid="&rsbody("bodyid")&" target='main'>"&rsbody("whname")&"</a><br>" & vbCrLf
           
		   
		   
       rsbody.movenext
	   loop 
	 end if 
	 rsbody.close
	 set rsbody=nothing
	%>
	
	</td>
    <td width="47">&nbsp;</td>
  </tr>
</table>
<br>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">设备管理系统 All Rights Reserved.</span>
  </tr>
</table>
</body>
</html>
<%



Call CloseConn
%>