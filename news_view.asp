<!--#include file="conn.asp"-->
<!--#include file="inc/imgcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<%
dim sql
dim rs
if request("id")<>"" then
	set rsnews=server.createobject("adodb.recordset")
	sqlnews="select * from xzgl_news where id="&request("id")
	rsnews.open sqlnews,conna,1,1
	title=rsnews("news_title")
end if
%>
<html>
<link href='css/index.css' rel='stylesheet' type='text/css'>
<head>
<title><%=title%>-信息管理系统</title>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<%
if request("action")="" then call main
if request("action")="savepost" then call savepost
if request("action")="del" then call del
sub savepost()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xzgl_news_re" 
      rsadd.open sqladd,connxzgl,1,3
      rsadd.addnew
      rsadd("body")=ReplaceBadChar(Trim(Request("body")))
      rsadd("news_id")=request("news_id")
      rsadd.update
      rsadd.close
	  'dwt.savesl conna.Execute("SELECT class_name FROM xzgl_news_class WHERE id="&request("news_class"))(0) ,"添加",ReplaceBadChar(Trim(Request("news_title")))
      set rsadd=nothing
	  dwt.out "<Script Language=Javascript>location.href='news_view.asp?id="&request("news_id")&"';</Script>"
end sub


 sub del()
dim rsdel2,sqldel2
ID=request("ID3")



	set rsdel2=server.createobject("adodb.recordset")
	sqldel="delete * from xzgl_news_re where id="&id
	rsdel2.open sqldel,connxzgl,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	'rsdel.close
	set rsdel2=nothing  
end sub


sub main()
'	set rsnews=server.createobject("adodb.recordset")
'	sqlnews="select * from xzgl_news where id="&request("id")
'	rsnews.open sqlnews,conna,1,1


%>
<!--#include file="index_t.asp"-->
<!--本站最新文章代码开始-->
<table width="760" border=0 align="center" cellpadding=0 cellspacing=0>
  <tr>
    <td class=main_title_1i>当前位置：<a href="/" class=class>信息管理系统首页</a>&gt;&gt;&gt; <a href="news_d.asp?NEWS_CLASS=<%=rsnews("news_class")%>"><%=newsclassh(rsnews("news_class"))%></a> &gt;&gt;&gt; <%=rsnews("news_title")%></td>
  </tr>
  <tr>
    <td  valign=top class=main_tdbg_575><br>
      <div align="center"><font color="#05006c" size=larger><%=rsnews("news_title")%></font> <br>
        <br>
        <hr width="80%" size="1">
      </div>
      <table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <%if rsnews("user_id")=0 then 
			   news_zz=rsnews("news_zz")
			else
				news_zz=usernameh(rsnews("user_id"))	
			end if    
				  %>
          <td align="center">发布时间：<font color="#990000"><%=rsnews("news_date")%></font>
            <%if rsnews("news_class")<>21 then%>
　 
发布：             <font color="#990000"><%=news_zz%> </font>
            <%end if%></td>
        </tr>
        <%
			  if not rsnews("isviewd") or session("userid")<>""  then 		 %>
        <tr>
          <td><br>
            <%
	response.write imgCode(rsnews("news_body"))
	Set rs = Server.CreateObject("adodb.recordset")
	sql = "select * from  xzgl_news where id="&request("id")
	rs.Open sql, Conna, 1, 3
	rs("view_numb") =rs("view_numb")+1      
	rs.Update
	rs.Close
	sql="SELECT isre FROM xzgl_news_class WHERE id="&rsnews("news_class")
	isre=conna.Execute(sql)(0)

    if isre then 
		dwt.out "<br/><br/><br/>当前1楼<br/> <hr width=100% size=1><br/><br/><br/>"
		post=1
		sqlpost="SELECT * FROM xzgl_news_re WHERE news_id="&rsnews("id")
		set rspost=server.createobject("adodb.recordset")
		rspost.open sqlpost,conna,1,1
		if rspost.eof and rspost.bof then 
			'dwt.out  message ("<p align='center'>未添加内容</p>" )
'		dwt.out "sdfsd"
		else
		   do while not rspost.eof
			post=post+1
			dwt.out rspost("body")
			dwt.out "<br/><br/><br/>当前"&post&"楼，"&rspost("date")&"<br/>"
if  session("groupid")=3  then 
				 dwt.out "&nbsp;&nbsp;&nbsp;<a href='?action=del&ID3="&rspost("id")&"' onClick=""return confirm('确定要删除此内容吗？');""><span style=""border-bottom-style: solid;border-width:2px;color:#660066"">删除</span></a>"
				 end if
dwt.out"<hr width=100% size=1><br/><br/><br/>"
		  rspost.movenext
		  loop
        end if
	end if  
    
	IF session("userid")<>"" AND rsnews("isviewd") THEN%>
            <DIV>
              <FIELDSET align=center>
              <LEGEND align=left>已浏览人员</LEGEND>
	<%if rsnews("isviewd") then 		
			if  rsnews("viewd")<>"" then  
			    V= Split(rsnews("viewd"),",") 
				For I = 0 To Ubound(V) 
				   if cint(V(I))=cint(session("userid")) then
						viewd=true
						exit FOR
				   end if 
				Next 
			end if 
			if not viewd then 
					set rsedit=server.createobject("adodb.recordset")
					sqledit="select * from xzgl_news where ID="&ReplaceBadChar(Trim(request("ID")))
					rsedit.open sqledit,connxzgl,1,3
					
						if  rsnews("viewd")<>"" then  
							rsedit("viewd")=rsedit("viewd")&","&session("userid")
						else
							rsedit("viewd")=session("userid")
						end if 
					rsedit.update
					rsedit.close
			end if 

'2008年10月16日改动，用于在签名分类where levelid<>10				 
'					sqluser="SELECT groupid FROM userid WHERE id="&session("userid")
'					usergroupid=conn.Execute(sqluser)(0)
'                    response.write usergroupid&"ooooooooooooooooo"
					
				 
	
	
	
				if  rsnews("viewd")<>"" then 
					dim sqlcj1,rscj1,record,bh,total
					if rsnews("viewdgroup")=1 then	sqlcj1="SELECT * from levelname where levelid=8"
					if rsnews("viewdgroup")=2 or rsnews("viewdgroup")=3  or isnull(rsnews("viewdgroup"))  then  sqlcj1="SELECT * from levelname "  '111220修改添加 OR ISNULL 当VIEWDGROUP为空时也输出人名
					set rscj1=server.createobject("adodb.recordset")
					rscj1.open sqlcj1,conn,1,1
					'record=rscj1.recordcount
					do while not rscj1.eof
				'		response.Write"<br>"
						totalgr=0
						response.Write"<br>"
						dwt.out "<STRONG>"&rscj1("levelname")&":</STRONG>&nbsp;&nbsp;&nbsp;"
						V= Split(rsnews("viewd"),",") 
						For I = 0 To Ubound(V) 
							dwt.out usernameh2(V(I),rsnews("viewdgroup"),rscj1("levelid"))
							'totalgr=totalgr+1
						Next
						
					response.Write("<b>总数:"&totalgr&"</b>")
					rscj1.movenext
					loop
				
					rscj1.close
					set rscj1=nothing  
				end if 
	 	   
'		for j =1 to record
'			totalgr=0
'			response.Write"<br>"
'			response.Write"<br>"
'			dwt.out "<STRONG>"&usernameh3(j)&":</STRONG>&nbsp;&nbsp;&nbsp;"
'		next
	end if 			%>
              </STRONG>
              </FIELDSET>
            </DIV>
            <%END IF %></td>
        </tr>
        <tr>
          <td align="right"><div align="right"><font color="#990000"><br>
              <br>
              【<a href="javascript:self.print()"><font color="#990000">打印该内容</font></a>】【<a href="javascript:window.close()"><font color="#990000">关闭窗口</font></a>】</font> </div>
            <%
'end sub
'sub addpost()
    if isre then 

    dwt.out "<div align=left> <hr width=100% size=1>"
	dwt.out "<form method='post' action='news_view.asp' name='form1' align=center>"
	dwt.out "   <textarea name='body' cols='40' rows='10'></textarea><br/>"	
	dwt.out"<input name='action' type='hidden' value='savepost'> <input name='news_id' type='hidden' value='"&rsnews("id")&"'>      <input  type='submit' name='Submit' value=' 回复 ' style='cursor:hand;'>"& vbCrLf
    dwt.out "</form></div>	"
end if    
   
'end sub	
%>
          </td>
        </tr>
        <%
			 
			 else
			 dwt.out "<tr><td>此内容要记录查看人员，请从首页登录后浏览</td></tr>"
			 end if 
			 
			 rsnews.close%>
      </table></td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
<!--外网搜索代码-->
<!--文章频道显示代码-->
<table width=760 border=0 align="center" cellpadding=0 cellspacing=0 
background=images2006/bottom_back.gif>
  <tbody>
    <tr>
      <td class=sxpta-font2 align=middle height=24>设备管理系统</td>
      <td width=140 height=54 rowspan=2><img height=54 
      src="images2006/bottom_r.gif" width=140 usemap=#Map 
  border=0></td>
    </tr>
    <tr>
      <td class=sxpta-font2 align=middle height=30><table class=black cellspacing=0 cellpadding=0 width=610 align=center 
      border=0>
          <tbody>
            <tr>
              <td width=170></td>
              <td valign=bottom width=394 height=28></td>
            </tr>
          </tbody>
        </table></td>
    </tr>
  </tbody>
</table>
<%end sub%>
</body>
</html>
