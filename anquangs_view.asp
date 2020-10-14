<!--#include file="conn.asp"-->
<!--#include file="inc/imgcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql
dim rs
if request("id")="" then
response.write "该程序执行了非法操作:)"
response.end
end if
set rspxst=server.createobject("adodb.recordset")
sqlpxst="select * from anquangs where id="&request("id")
rspxst.open sqlpxst,connaq,1,1
  

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=rspxst("pxst_title")%></title>
<link title="css" href="../css2012/index.css" rel="stylesheet"  type="text/css"/>
<LINK href="../css2012/menu.css" type=text/css rel=stylesheet>
<SCRIPT src="../css2012/menu.js" type=text/javascript></SCRIPT>
<SCRIPT language=javascript src="js/hhh.js"></SCRIPT>
<script language=javascript>
   <!--
    function CheckForm() {
      if(document.Login.UserName.value == '') {
        alert('请输入用户名！');
        document.Login.UserName.focus();
        return false;
      }
      if(document.Login.password.value == '') {
        alert('请输入密码！');
        document.Login.password.focus();
        return false;
      }
	  }
  //-->

</script>
</head>

<body>
<!--#include file="top.asp"-->

<DIV class="box960">
<div class="boxl">
<div class="t1">
  <div class="dq"><a href="/" class=class>首页</a> &gt; <a href=anquangs_d.asp?wangong=0>安全短板公示</a> </div>
</div>
<p class="br"> </p>
<div class="center boxlc boxlt">
<%if rspxst("userid")=0 then 
			   news_zz=rspxst("pxst_zz")
			else
				news_zz=usernameh(rspxst("userid"))	
			end if    
				  %>

发布时间：<font color="#990000"><%=rspxst("pxst_date")%></font>　 任务发布人：<font color="#990000"><%=news_zz%> </font>
<hr>
<div class='bodyt'><br>
  <strong>提出时间：</strong><%=rspxst("huiyi_date")%><br>
  <strong>责任单位：</strong><%=rspxst("zr_danwei")%><br>
  <strong>责任人：</strong><%=rspxst("zr_ren")%><br>
  <strong>要求完工时间：</strong><%=rspxst("yaoqiu_date")%><br>
  <strong>实际完工时间：</strong><%=rspxst("wangong_date")%></div>
<div class='bodyt'><strong>安全短板内容及进度描述如下：</strong><br><%=rspxst("pxst_title")%><%'imgCode(rspxst("pxst_body"))%></div>
<br>
<%set rsqxtb_fk=server.createobject("adodb.recordset")
sqlqxtb_fk="select * from anquangs_fk where huiyiluoshi_id="&request("id")
rsqxtb_fk.open sqlqxtb_fk,connaq,1,1
if rsqxtb_fk.eof and rsqxtb_fk.bof then 
response.write "<p align='center'>暂无反馈</p>" 
else
do while not rsqxtb_fk.eof 
i=i+1


if i mod 2= 0 then %>
<table width="80%"  border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#3399CC">
<%else
			  %>
<table width="80%"  border="1" align="center" cellpadding="0" cellspacing="0">
  <%end if %>
  <tr>
    <td align="center">反馈时间：<font color="#990000"><%=rsqxtb_fk("huiyiluoshi_fk_date")%></font>　 反馈：<font color="#990000"><%=sscjh(rsqxtb_fk("huiyiluoshi_fk_sscj"))%> </font></td>
  </tr>
  <tr>
    <td><br>
      <%=rsqxtb_fk("huiyiluoshi_fk_body")%></td>
  </tr>
</table>
<%
			    rsqxtb_fk.movenext
          loop
       end if
	   rsqxtb_fk.close
set rsqxtb_fk=nothing%>
<div align="right"><font color="#990000"> 【<a href="javascript:self.print()"><font color="#990000">打印该内容</font></a>】【<a href="javascript:window.close()"><font color="#990000">关闭窗口</font></a>】</font> </div>
<br>
<br>
</div>
</div>

<!--#include file="index_left.asp"-->

</div>
<div class="clear"></div>
<div class=miniNav>
  <div class="box960" align="center"><br>
    <br>
    <b>设备管理系统</b> <br>
    <br>
  </div>
</div>
</body>
</html>
