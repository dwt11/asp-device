<!--#include file="conn.asp"-->
<!--#include file="inc/imgcode.asp"-->
<!--#include file="inc/function.asp"-->


<%
dim sql
dim rs
'if request("id")="" then
'response.write "该程序执行了非法操作:)"
'response.end
'end if
set rsqxdj=server.createobject("adodb.recordset")
sqlqxdj="select * from mzqxdjzg where jcdate1=#"&request("jcdate1")&"# and jcdate2=#"&request("jcdate2")&"#"
rsqxdj.open sqlqxdj,connb,1,1
%>
<html>
<link href='/DefaultSkin.css' rel='stylesheet' type='text/css'> 

<head>
<title><%dwt.out request("jcdate1")&"到"&request("jcdate2")%>-每周缺陷记录</title>
<link href='css/ext-all.css' rel='stylesheet' type='text/css'>
<link href='css/body.css' rel='stylesheet' type='text/css'>

<link rel="stylesheet" href="../style/style.css" type="text/css">
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
            <td class=main_title_1i>当前位置：<a href="/" class=class>信息管理系统首页</a> &gt;&gt;&gt; 每周缺陷记录&gt;&gt;&gt; <%dwt.out request("jcdate1")&"到"&request("jcdate2")&"未整改内容"%></td>
          </tr>
          
          <tr>
            <td  vAlign=top class=main_tdbg_575><br><div align="center"><font color="#05006c" size=larger><%dwt.out request("jcdate1")&"到"&request("jcdate2")&"未整改内容"%></font> <br>
              <br><hr width="80%" size="1">
              </div>
              <table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td>
<%
	if rsqxdj.eof and rsqxdj.bof then 
	   message "未找到相关内容"
	else
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>位号</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>缺陷内容</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>检查日期</div></td>"
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>整改状态</div></td>"
		dwt.out "    </tr>"


do while not rsqxdj.eof 
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&xh_id&"</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"
			dwt.out sscjh_d(rsqxdj("sscj"))&"</div></td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"""
			'if now()-rsqxdj("update")<7 then   dwt.out "bgcolor=""#FFFF00"""
			dwt.out ">"
			if rsqxdj("zgzt") then 
			   dwt.out searchH(uCase(rsqxdj("wh")),keys)
			else
			   dwt.out "<font color='#ff0000'>"&searchH(uCase(rsqxdj("wh")),keys)&"<font>"
			end if  
			   dwt.out "&nbsp;</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rsqxdj("body")&"&nbsp;</td>"
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&rsqxdj("jcdate1")&"到"&rsqxdj("jcdate2")&"</div></td>"
			if rsqxdj("zgzt") then 
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">已整改</td>"
			else
			   dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">未整改</td>"
			end if 
			dwt.out "</tr>"
			    rsqxdj.movenext
          loop

		dwt.out "</table>"
end if
                      %> </td>
                </tr>
                <tr>
                  <td align="right"><div align="right"><font color="#990000"><br><br>【<a href="javascript:self.print()"><font color="#990000">打印该内容</font></a>】【<a href="javascript:window.close()"><font color="#990000">关闭窗口</font></a>】</font> </div></td>
                </tr>
              </table>
			  
			  <%rsqxdj.close
set rsqxdj=nothing%>

			  
			</td>
          </tr>
</table>
<TABLE width=760 border=0 align="center" cellPadding=0 cellSpacing=0 
background=images2006/bottom_back.gif>
  <TBODY>
  <TR>
    <TD class=sxpta-font2 align=middle height=24>设备管理系统</TD>
    <TD width=140 height=54 rowSpan=2><IMG height=54 src="images2006/bottom_r.gif" width=140   border=0></TD></TR>
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


