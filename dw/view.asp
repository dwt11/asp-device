<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="function.asp"-->
<%
bodyid=request("id")
if bodyid<>"" then
	set rsnews=server.createobject("adodb.recordset")
	sqlnews="select * from dgtzl_body where id="&request("id")
	rsnews.open sqlnews,conndgt,1,1
	title=rsnews("news_title")
else
 response.End
end if  



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=title%>-党建-信息管理系统</title>
<link title="css" href="../css2012/index.css" rel="stylesheet"  type="text/css"/>
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
      <div class="dq">当前位置：<a href="/">首页</a> > <a href="/dw">党建</a> <%=gettclassname(rsnews("index"))%> > <%=getclassname(rsnews("index"))%> </div>
    </div>
    <p class="br"> </p>
    <div class="center boxlc boxlt">
      <h1><%=rsnews("news_title")%></h1>
      <%if rsnews("user_id")=0 then 
			   news_zz=rsnews("news_zz")
			else
				news_zz=usernameh(rsnews("user_id"))	
			end if    
				  %>
      日期： <%=gettdate(rsnews("news_date"))%> &#160;&#160;&#160;&#160;浏览次数: <%=rsnews("view_numb")%>次&#160;&#160;&#160;&#160;发布：             <font color="#990000"><%=news_zz%> </font>
      <hr/>
      
        <%
			  if not rsnews("isviewd") or session("userid")<>""  then 		
			  
		response.write "<div class='bodyt'>"& imgCode(rsnews("news_body"))&"      </div>"
 	Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from  dgtzl_body where id="&request("id")
    rs.Open sql, Conndgt, 1, 3
	rs("view_numb") =rs("view_numb")+1      
	rs.Update
	rs.Close
	sql="SELECT isre FROM dgtzl_class WHERE id="&rsnews("news_class")
	isre=conndgt.Execute(sql)(0)
	  
			  
			  
			  
			  
			  
		IF session("userid")<>"" AND rsnews("isviewd") THEN%>
<DIV class=fieldset>
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
					sqledit="select * from dgtzl_body where ID="&ReplaceBadChar(Trim(request("ID")))
					rsedit.open sqledit,conndgt,1,3
					
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
					if rsnews("viewdgroup")=2 or rsnews("viewdgroup")=3 or isnull(rsnews("viewdgroup")) or rsnews("viewdgroup")=0 then  sqlcj1="SELECT * from levelname "
					set rscj1=server.createobject("adodb.recordset")
					rscj1.open sqlcj1,conn,1,1
					'record=rscj1.recordcount
					do while not rscj1.eof
				'		response.Write"<br>"
						totalgr=0
						dwt.out "<STRONG>&nbsp;&nbsp;&nbsp;"&rscj1("levelname")&":</STRONG>&nbsp;&nbsp;&nbsp;"
						V= Split(rsnews("viewd"),",") 
						For I = 0 To Ubound(V) 
							dwt.out usernameh2(V(I),rsnews("viewdgroup"),rscj1("levelid"))
							'totalgr=totalgr+1
						Next
						
					response.Write("<b>总数:"&totalgr&"</b><br/>")
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
             </FIELDSET>
             </DIV>
            <%END IF %>
        <div align="right"><font color="#990000">
          【<a href="javascript:self.print()"><font color="#990000">打印该内容</font></a>】【<a href="javascript:window.close()"><font color="#990000">关闭窗口</font></a>】</font> </div>
        <%
			 
			 else
			 dwt.out "<tr><td>此内容要记录查看人员，请从首页登录后浏览</td></tr>"
			 end if 
			 
			 rsnews.close%>
    </div>
  </div>
  
  <!--#include file="left.asp"--> 
  
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
