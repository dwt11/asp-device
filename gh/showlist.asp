<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="function.asp"-->
<%
classid=request("classid")
if classid<>"" then classname=getclassname(classid)

      url="showlist.asp?classid="&classid


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=classname%>-工会-信息管理系统</title>
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
      <div class="dq">当前位置：<a href="/">首页</a> > <a href="/gh">工会</a> <%=gettclassname(classid)%> > <%=classname%> </div>
    </div>
    <p class="br"> </p>
    <div class="boxlc boxlt">
      <%			
				sqltree3="SELECT * from dgtzl_body where index="&classid&" and news_class=2 order by id desc"& vbCrLf
				set rstree3=server.createobject("adodb.recordset")
				rstree3.open sqltree3,conndgt,1,1
				if  rstree3.eof then 
					dwt.out "&nbsp;&nbsp;&nbsp;暂无内容"
				else
%>
      <div class="list">
        <%

				   record=rstree3.recordcount
				   if Trim(Request("PgSz"))="" then
					   PgSz=20
				   ELSE 
					   PgSz=Trim(Request("PgSz"))
				   end if 
				   rstree3.PageSize = Cint(PgSz) 
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
				   rstree3.absolutePage = page
				   start=PgSz*Page-PgSz+1
				   rowCount = rstree3.PageSize
				
					do while not rstree3.eof and rowcount>0
						title=rstree3("news_title")
						if len(title)>30 then
						title=left(title,28)&"..."
						end if 
						dwt.out "<li><a href=view.asp?id="&rstree3("id")&" title='"&rstree3("news_title")&"'>"&title&"</a> <span>["&gettdate(rstree3("news_date"))&"]</span></li>"
			
			                 RowCount=RowCount-1

					rstree3.movenext
					loop
        %>
      </div>
      <%
				
call newshowpage(page,url,total,record,PgSz)
				end if 				
				
				
				
				

		%>
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
