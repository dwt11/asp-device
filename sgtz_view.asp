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
set rsnews=server.createobject("adodb.recordset")
sqlnews="select * from sgtz where id="&request("id")
rsnews.open sqlnews,connb,1,1
%>
<html>
<link href='css/index.css' rel='stylesheet' type='text/css'> 

<head>
<title><%=rsnews("wh")%>-信息管理系统</title>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
</TR>
</TBODY>
</TABLE>
</TD>
<TD></TD>
    </TR>
  </TBODY>
</TABLE>



  <table width="760" height="120" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td background="/images2006/logo.jpg">&nbsp;</td>
    </tr>
</table>
  
  <TABLE width=760 
  border=0 align="center" cellPadding=0 cellSpacing=0 class=table-left1-right2>
    <TBODY>
  <TR>
    <TD vAlign=bottom>
      <TABLE height=32 cellSpacing=0 cellPadding=0 width="100%" 
      background=images2006/topmenu-bg.gif border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE class=sxpta-font1 cellSpacing=0 cellPadding=0 width=750 
            border=0>
              <TBODY>
              <TR>
                <TD align=right width=41 
                background=images2006/topmenu-bg2.gif><IMG height=11 
                  src="images2006/d.gif" width=13 name=Image2></TD>
                <TD width=238 
                  background=images2006/topmenu-bg2.gif><div align="center"><A class=link2 
                  href="/index.asp">首页</A></div></TD>
                <TD align=middle width=15><IMG height=32 
                  src="images2006/menuicon2.gif" width=15></TD>
                <TD align=middle width=341>新闻</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>技术资料</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>管理规程</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>值班表</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>通知</TD>
                </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
  <table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
    <tr>
      <td vAlign=top>
      <!--本站最新文章代码开始-->
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td class=main_title_1i>当前位置：<a href="/" class=class>信息管理系统首页</a> &gt;&gt;&gt;  <%=rsnews("wh")%></td>
          </tr>
          <tr>
            <td  vAlign=top class=main_tdbg_575><br><div align="center"><font color="#05006c" size=larger><%=rsnews("wh")%></font> <br>
              <br><hr width="80%" size="1">
              </div>
              <table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
<!--                <tr>
                  <%'if rsnews("user_id")=0 then 
				          ' news_zz=rsnews("news_zz")
					'else
					'    news_zz=usernameh(rsnews("user_id"))	
					'end if    
				  %>
				  
				  <td align="center">发布时间：<font color="#990000"></font>　 发布：<font color="#990000"> </font></td>
                </tr>
-->               
                 <tr >
                  <td align="right"> 事故位号名称：</td>
                  <td align="left" width="80%"> <%=rsnews("wh")%></td>
                </tr>
                 <tr >
                  <td align="right"> 事故地点：</td>
                  <td align="left" width="80%"> <%=rsnews("address")%></td>
                </tr>
                 <tr >
                  <td align="right"> 事故时间：</td>
                  <td align="left" width="80%"> <%=rsnews("createdate")%> </td>
                </tr>
                 <%			if rsnews("class")=1 then sgclass="设备事故"
			if rsnews("class")=2 then sgclass="操作事故"
			if rsnews("class")=3 then sgclass="责任事故"

				  %>
				  
                 <tr >
                  <td align="right"> 事故类别：</td>
                  <td align="left"> <%=sgclass%> </td>
                </tr>
                
                 <tr >
                  <td align="right"> 责任人：</td>
                  <td align="left"> <%=rsnews("ren")%> </td>
                </tr>
                 <tr >
                  <td align="right"> 责任人处理：</td>
                  <td align="left"> <%=rsnews("cljg")%> </td>
                </tr>
                 <tr >
                  <td align="right"> 事故经过及主要原因：</td>
                  <td align="left"> <%=rsnews("jg")%> </td>
                </tr>
                 <tr >
                  <td align="right"> 防范措施：</td>
                  <td align="left"> <%=rsnews("clcs")%> </td>
                </tr>
                 <tr >
                  <td align="right"> 备注：</td>
                  <td align="left"> <%=rsnews("bz")%> </td>
                </tr>
                
                
                
                
                
                
                <tr>
                  <td></td>
                  <td align="right"><div align="right"><font color="#990000"><br><br>【<a href="javascript:self.print()"><font color="#990000">打印该内容</font></a>】【<a href="javascript:window.close()"><font color="#990000">关闭窗口</font></a>】</font> </div></td>
                </tr>
              </table>
			</td>
          </tr>
          <tr>
            <td ></td>
          </tr>
        </table>
      </td>
    </tr>
</table>
  <!--外网搜索代码-->
  <!--文章频道显示代码-->
<TABLE width=760 border=0 align="center" cellPadding=0 cellSpacing=0 
background=images2006/bottom_back.gif>
  <TBODY>
  <TR>
    <TD class=sxpta-font2 align=middle height=24>设备管理系统</TD>
    <TD width=140 height=54 rowSpan=2><IMG height=54 
      src="images2006/bottom_r.gif" width=140 useMap=#Map 
  border=0></TD></TR>
  <TR>
    <TD class=sxpta-font2 align=middle height=30>
      <TABLE class=black cellSpacing=0 cellPadding=0 width=610 align=center 
      border=0>
        <TBODY>
        <TR>
          <TD width=170>
            </TD>
          <TD vAlign=bottom width=394 height=28>
</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>

</body>
</html>


