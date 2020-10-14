<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<html>
<head>
<title>缺陷整改通知</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href='/css/index.css' rel='stylesheet' type='text/css'> 
<script language='JavaScript' type='text/JavaScript' src='/js/menu.js'></script>
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
<body leftmargin=0 topmargin=0  onmousemove='HideMenu()'>

<!--#include file="index_t.asp"-->

</TABLE> <table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
    <tr>
      <td width=180 vAlign=top>
      <!--用户登录代码开始-->
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td><IMG src="/images2006/login_01.gif"></td>
          </tr>
          <tr>
            <td vAlign=center align=middle background=/images2006/login_02.gif><form name='Login' action='login.asp' method='post' target='_parent'  onSubmit='return CheckForm();'>
<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'>
            <tr>
                <td height='25' align='right'>用户名：</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td>
       
                </tr>
                <tr>
     
                <td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='password' type='password' id='Password' size='10' maxlength='20'></td>
      
                </tr>
                <tr align='center'>
                  <td height='47' colspan='2'>
       
                           <input type='hidden' name='Action' value='Login'>
		  <input type="submit" name="Submit" value="登录">
&nbsp;&nbsp;<input name='Reset' type='reset' id='Reset' value=' 清除 '>
 </td>
        </tr>
		</table>
            </form>
        </td>
          </tr>
          <tr>
            <td><IMG src="/images2006/login_03.gif">
              <table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=left_title align=middle>缺陷整改通知</td>
                </tr>
                <tr>
                  <td class=left_tdbg1 vAlign=top height=179><%
sqlqxtb="SELECT top 8 * from scgl_qxtb ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connb,1,1
if rsqxtb.eof and rsqxtb.bof then 
response.write "<p align='center'>未添加新闻</p>" 
else
do while not rsqxtb.eof

title=rsqxtb("qxtb_title")
if len (title)>25 then
title=left(title,25)&"..."

%>
                      
                        <li><a href="qxtb_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("qxtb_title")%>" target=_blank><%=title%></a><br>
                            <%else%>
                        <li><a href="qxtb_view.asp?id=<%=rsqxtb("id")%>" target=_blank><%=rsqxtb("qxtb_title")%></a><br>
                            <%end if%>
                            <%i=i+1
if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
%>
                    </td>
                </tr>
                <tr>
                  <td class=left_tdbg2></td>
                </tr>
              </table>
              <table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=left_title align=middle>通知</td>
                </tr>
                <tr>
                  <td class=left_tdbg1 vAlign=top height=120>
                    <table width='100%' cellpadding='0' cellspacing='5' border='0' align='center'>
                      <tr valign='top'>
                        <td align='center'><br>
                        </td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td class=left_tdbg2></td>
                </tr>
              </table>
              <table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=left_title align=middle>管理规定</td>
                </tr>
                <tr>
                  <td class=left_tdbg1 vAlign=top height=126>&nbsp;</td>
                </tr>
                <tr>
                  <td class=left_tdbg2></td>
                </tr>
              </table></td>
          </tr>
        </table>
      <!--用户登录代码结束--></td>
      <td width=5></td>
      <td vAlign=top>
	  <%
	  
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><div align=""center""><strong>缺陷整改通知标题</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>发布者</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>发布时间</strong></div></td>"
response.write "    </tr>"
sqlqxtb="SELECT * from scgl_qxtb ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connb,1,1
if rsqxtb.eof and rsqxtb.bof then 
response.write "<p align='center'>未添加缺陷整改通知</p>" 
else
           record=rsqxtb.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsqxtb.PageSize = Cint(PgSz) 
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
           rsqxtb.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsqxtb.PageSize
           do while not rsqxtb.eof and rowcount>0
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rsqxtb("id")&"</div></td>"
                 response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=qxtb_view.asp?id="&rsqxtb("id")&" target=_blank>"&rsqxtb("qxtb_title")&"</a></td>"
                 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&usernameh(rsqxtb("userid"))&"</div></td>"
				 
				 
				 response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqxtb("qxtb_date")&"</div></td>"
                 response.write "    </tr>"
                 RowCount=RowCount-1
          rsqxtb.movenext
          loop
       end if
       rsqxtb.close
       set rsqxtb=nothing
        conn.close
        set conn=nothing
        response.write "</table>"
       call showpage1(page,url,total,record,PgSz)
	  %>
	  
	  
	  
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
  border=0></TD>
      </TR>
      <TR>
        <TD class=sxpta-font2 align=middle height=30>
          <TABLE class=black cellSpacing=0 cellPadding=0 width=610 align=center 
      border=0>
            <TBODY>
              <TR>
                <TD width=170> </TD>
                <TD vAlign=bottom width=394 height=28> </TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      </TR>
    </TBODY>
  </TABLE>
</body>
</html>