<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<html>
<head>
<%
if request("news_class")="" then
 title="信息管理系统"
else
 title=newsclassh(request("news_class"))&"-信息管理系统"
end if 
%>

<title><%=title%>-</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href='css/index.css' rel='stylesheet' type='text/css'> 
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
<body leftmargin=0 topmargin=0  >
 
	<!--#include file="index_t.asp"-->
			
				<table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
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
if request("news_class")="" then 
  sqlqxtb="SELECT top 8 * from scgl_qxtb ORDER BY id DESC"
else
  sqlqxtb="SELECT top 8 * from scgl_qxtb ORDER BY id DESC"
end if   
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connb,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.Out "<p align='center'>未添加内容</p>" 
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
            </td>
          </tr>
        </table>
      <!--用户登录代码结束--></td>
      <td width=5></td>
      <td vAlign=top>
	  <%
	  
	  Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
Dwt.Out "<tr class=""title"">" 
Dwt.Out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center""><strong>序号</strong></Div></td>"
Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center""><strong>标题</strong></Div></td>"
Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center""><strong>分类</strong></Div></td>"
Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>发布者</strong></Div></td>"
Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>发布时间</strong></Div></td>"
Dwt.Out "    </tr>"
if request("news_class")="" then 
    sqlnews="SELECT xzgl_news.*,xzgl_news_class.isindex FROM xzgl_news INNER JOIN xzgl_news_class ON xzgl_news.news_class=xzgl_news_class.id WHERE (((xzgl_news_class.isindex)=True))  ORDER BY xzgl_news.news_date desc"
    url="news_d.asp"
else
    sqlnews="SELECT * FROM xzgl_news where news_class="&request("news_class")&" ORDER BY id DESC"
      url="news_d.asp?news_class="&request("news_class")
end if 	
set rsnews=server.createobject("adodb.recordset")
rsnews.open sqlnews,conna,1,1
if rsnews.eof and rsnews.bof then 
Dwt.Out "<p align='center'>未添加新闻</p>" 
else
           record=rsnews.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsnews.PageSize = Cint(PgSz) 
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
           rsnews.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsnews.PageSize
           do while not rsnews.eof and rowcount>0
                 xh=xh+1
                 Dwt.Out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 Dwt.Out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"</Div></td>"
                 Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px"" ><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</a></td>"
                 Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px"" ><a href='news_d.asp?NEWS_CLASS="&rsnews("news_class")&"'>"&newsclassh(rsnews("news_class"))&"</a></td>"
                 Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
if rsnews("user_id")=0 then 
				Dwt.out rsnews("news_zz")
			  else
				Dwt.out usernameh(rsnews("user_id"))
			  end if 
Dwt.out"</Div></td>"
                 Dwt.Out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&year(rsnews("news_date"))&"-"&month(rsnews("news_date"))&"-"&day(rsnews("news_date"))&"</Div></td>"
                 Dwt.Out "    </tr>"
                 RowCount=RowCount-1
          rsnews.movenext
          loop
       end if
       rsnews.close
       set rsnews=nothing
        conn.close
        set conn=nothing
        Dwt.Out "</table>"
if request("news_class")="" then        
   call showpage1(page,url,total,record,PgSz)
else 
call showpage(page,url,total,record,PgSz)
	end if   
	  
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