<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="function.asp"-->
<%
classid=request("classid")
if classid<>"" then classname=getclassname(classid)&"-"



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=classname%>工会-信息管理系统</title>
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
      <div class="dq">当前位置：<a href="/">首页</a> > <a href="/gh">工会</a>  <%=gettclassname(classid)%> </div>
    </div>
    <%dim ii
		if classid="" then sqltree="SELECT * from dgtzl_index_gh where index=0 order by orderby asc"
		if classid<>"" then sqltree="SELECT * from dgtzl_index_gh where index="&classid&" order by orderby asc"
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,conndgt,1,1
		
		
		do while not rstree.eof
			dim urltmp1,titlecss,classid1
			ii=ii+1
			if ii mod 2 =1 then titlecss="boxl_1"
			if ii mod 2 =0 then titlecss="boxl_2"
			
			
			
			'获取子栏目
			
			sqltree2="SELECT * from dgtzl_index_gh where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree2=server.createobject("adodb.recordset")
			rstree2.open sqltree2,conndgt,1,1
			if rstree2.eof then 
			    classid1=rstree("id")
				urltmp="showlist.asp?classid="&rstree("id")
			else
			    classid1=""
				do while not rstree2.eof
					classid1=classid1&rstree2("id")&","
				rstree2.movenext
			    urltmp="index.asp?classid="&rstree("id")
				loop
			end if 
			
		
	
			'dwt.out "<br>"
			
				dwt.out "<div class='"&titlecss&" boxlc'>"
				dwt.out "  <div class='blt'><a href="&urltmp&"><SPAN>更多...</SPAN>"&rstree("class_name")&"</a></div>"
				dwt.out "  <div class='blc' style='height:180px'>"
			
				sqltree3="SELECT top 7 * from dgtzl_body where index in("&classid1&") and news_class=2 order by id desc"& vbCrLf
				set rstree3=server.createobject("adodb.recordset")
				rstree3.open sqltree3,conndgt,1,1
				if  rstree3.eof then 
					dwt.out "暂无内容"
				else

					
					do while not rstree3.eof
						title=rstree3("news_title")
						if len(title)>15 then
						title=left(title,14)&"..."
						end if 
						dwt.out "<li><a href=view.asp?id="&rstree3("id")&" title='"&rstree3("news_title")&"'>"&title&"</a> <span>["&getsdate(rstree3("news_date"))&"]</span></li>"
					rstree3.movenext
					loop
				end if 				
				
				
				
				
				
				
				 dwt.out " </div>"
				  dwt.out "<div class='blb'></div>"
				dwt.out "</div>"
			
			
			
			
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>
    
    
    
    
    
    
    
    
    
    
    
   
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
