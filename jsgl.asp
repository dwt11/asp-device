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
<title>系统管理首页</title>
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
    <td height="34" ><div align="center" class="style2" >技术管理首页</div></td>
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
    <td height="50" >技术管理包括以下功能：<a href="ylb.asp">技术档案</a>&nbsp;&nbsp;<a href="lsda.asp">联锁档案</a>&nbsp;&nbsp;<a href="fdbw.asp">防冻保温</a>
技改台账&nbsp;&nbsp;资料图纸&nbsp;&nbsp;检修记录&nbsp;&nbsp;设备运转率&nbsp;&nbsp;培训管理 </td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;1、技术档案 全厂设备总括，分为：<a href="ylb.asp?lxclassid=1">电接点压力表</a> <a href="ylb.asp?lxclassid=2">变送器</a> <a href="ylb.asp?lxclassid=3">转换器</a> <a href="ylb.asp?lxclassid=4">调节阀附件</a> <a href="ylb.asp?lxclassid=5">电磁阀</a> <a href="ylb.asp?lxclassid=6">就地调节器</a> <a href="ylb.asp?lxclassid=7">转速探头</a> <a href="ylb.asp?lxclassid=8">流量一次元件</a> <a href="ylb.asp?lxclassid=9">测温一次元件</a> <a href="ylb.asp?lxclassid=10">机组探头</a> <a href="ylb.asp?lxclassid=11">分析</a> <a href="ylb.asp?lxclassid=12">空调</a> <a href="ylb.asp?lxclassid=13">皮带秤</a> <a href="ylb.asp?lxclassid=14">调节阀</a> <a href="ylb.asp?lxclassid=15">电动执行机构</a>十五个分类。</td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;2、联锁档案 全厂联锁统计汇总。</td>
  </tr>
  <tr>
    <td height="28" >&nbsp;&nbsp;&nbsp;&nbsp;3、防冻保温 全厂冬季防冻保温表统计汇总。</td>
  </tr>
</table>

</body>
</html>
<%



Call CloseConn
%>