<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<html Xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>信息管理系统</title>
<LINK href="css/docs.css" type="text/css" rel="stylesheet"></LINK>
<LINK href="css/ext-all.css" type="text/css" rel="stylesheet">
<script src="js/info.js"></script>

<style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style><SCRIPT>
function BarMove(){
 if (AtMovePic2.style.display==""){
  document.all("AtMovePic2").style.display="none"
  document.all("AtMovePic").style.display=""
  document.all("frmTitle").style.display="none"
 }
 else{
  document.all("AtMovePic2").style.display=""
  document.all("AtMovePic").style.display="none"
  document.all("frmTitle").style.display=""
 }
}
</SCRIPT>

<script>
<!--
$(function(){
	$(window).load(function(){
		$("div[id=newnotice]").css({"right":"0px","bottom":"1px"});
		$("div[id=newnotice]").slideDown("slow");
		
		/*setTimeout(function(){$("div[id=newnotice]").slideUp("slow")},10000);*/
	}).scroll(function(){
		$("div[id=newnotice]").css({"bottom":"0px"});
		$("div[id=newnotice]").css({"right":"0px","bottom":"1px"});
	}).resize(function(){
		$("div[id=newnotice]").css({"bottom":""});
		$("div[id=newnotice]").css({"right":"0px","bottom":"1px"});	
	});
	
	$("label[id=tomin]").click(function(){
		$("div[id=noticecon]","div[id=newnotice]").slideUp();
	});
	
	$("label[id=tomax]").click(function(){
		$("div[id=noticecon]","div[id=newnotice]").slideDown();
	});
	
	$("label[id=toclose]").click(function(){
		$("div[id=newnotice]").hide();
	});
});
//scroll : 滚动时候保持在页面右侧底部.
//resize: 浏览器变化时候  保持在页面右侧底部.
-->
</script>
<style>
<!--
#newnotice {
	position:absolute;
	display:none;
	width:250px;
	/*height:22px;*/
	border:solid #9CBCE8 1px;
	background-color: #F0FBEB
}
#newnotice p {
	font-size:12px;
	margin:1px;
	padding:0px 2px 0px 5px;
	background-color:#D9E5FA;
	color:#666666;
	height:20px;
	line-height:20px;
}
#newnotice p .title {
	float:left;
}
#newnotice p #bts {
	display:block;
	float:right;
	width:48px;
	height:15px;
	/*border:#000000 solid 1px;*/
}
#newnotice p #bts .button {
	display:block;
	float:left;
	width:15px;
	height:15px;
	line-height:15px;
	cursor:pointer;
	/*border:#000000 solid 1px;*/
}
#newnotice p #bts #tomin {
	background-image:url(img_ext/notice_button.gif);
	background-position:center;
}
#newnotice p #bts #tomax {
	background-image:url(img_ext/notice_button.gif);
	background-position:bottom;
}
#newnotice p #bts #toclose {
	background-image:url(img_ext/notice_button.gif);
}
#newnotice div {
	font-size:12px;
	margin:1px;
	padding:0px 5px 0px 5px;
	background-color:#FFFFFF;
	color:#999999;
	height:75px;
	line-height:20px;
}
-->
</style>

</head>

<body class=" ext-ie x-layout-container " id="docs" scroll="no" style="margin:0px;">
<DIV class="x-layout-panel x-layout-panel-north" id="ext-gen6" style="WIDTH: 100%; HEIGHT: 29px">
    <DIV class=" x-layout-active-content" id="header">
		<DIV style="FLOAT: right;PADDING-TOP: 5px;color:#ffffff;font-size:12px" ><a href="main.asp"   style="color:ffffff;">首页</a>&nbsp;&nbsp;<a href="login.asp?action=Logout" style="color:ffffff;">退出</a>&nbsp;&nbsp;&nbsp;&nbsp; 今天是<span id="webasp_time"></span><script>setInterval("webasp_time.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt (new Date().getDay());",1000);</script></div>
		 <DIV style="PADDING-TOP: 3px">   &nbsp;&nbsp;&nbsp;&nbsp;信息管理系统</DIV>
    </DIV>
</DIV>
<div style="padding-top:30px;">
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%" style="	padding-TOP: 30px;">
  <TBODY>
    <TR >
      <TD  align=middle id=frmTitle noWrap rowSpan=3 vAlign=center name="frmTitle">
       
	    <table width="100%" height="100%" border=1 cellpadding=0 cellspacing=0 bordercolor="#98c0f4" style="border-collapse:collapse"> 
          <tr>
            <td height="100%" colspan="2" align=middle><iframe frameborder=0 id=menu name=menu src="left.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 176px; Z-INDEX: 2"></iframe></td>
          </tr>
        </table>
		
    <TD rowSpan=3>
      
	  <TABLE width="6" height="100%" cellPadding=0 cellSpacing=0 bgcolor="#FF5E00">
        <TBODY>
          <TR>
            <TD width="1" class="x-layout-split x-layout-split-west x-splitbar-h x-layout-split-h" vAlign=top id=AtMovePic style=" WIDTH: 6px;HEIGHT: 640px;display:none;CURSOR: hand" onclick=BarMove() name="AtMovePic">              </TD>
              <TD width="1" rowspan="2" vAlign=top class="x-layout-split x-layout-split-west x-splitbar-h x-layout-split-h" id=AtMovePic2 style="top:2px;HEIGHT: 640px; WIDTH: 6px;CURSOR: hand" onclick=BarMove() name="AtMovePic2">              </TD>
      </TR>
          </TBODY>
        </TABLE>
	  <TD style="HEIGHT: 100%">
      <TD style="WIDTH: 100%">
  
  <TABLE width="100%" height="95%" cellPadding=0 cellSpacing=0  border=1 bordercolor="#98c0f4" style="border-collapse:collapse">
   <TBODY>
   <TR>
     <TD height="100%" colspan="2" align=middle>
     <IFRAME frameBorder=0 id=main scrolling="AUTO" name=main src="right.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></IFRAME></TD>
   </TR>
    </TBODY>
  </TABLE>  

   </td>
    </TR>
 </TBODY>
</TABLE>
<%
dim numb,isviewd
numb=0
set rsnews=server.createobject("adodb.recordset")
sqlnews="select * from xzgl_news where isviewd ORDER BY id desc"
rsnews.open sqlnews,conna,1,1
if rsnews.eof and rsnews.bof then 
	
else
	do while not rsnews.eof
         isviewd=false
		
sqluser="SELECT groupid FROM userid WHERE id="&session("userid")
usergroupid=conn.Execute(sqluser)(0)
	if rsnews("viewdgroup")=1 and usergroupid=10 then
		if  rsnews("viewd")<>"" then  
		 V= Split(rsnews("viewd"),",") 
			For I = 0 To Ubound(V) 
			   if cint(V(I))=cint(session("userid")) then 
				   isviewd=true
				   'info=info&cint(V(I))&"-----"&Ubound(V)&"-"&I
				   exit for    
			   end if 	   
			Next 
			if isviewd=false then 
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
			end if 
		else
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
		end if 
	end if 
	
	
	if rsnews("viewdgroup")=2 and (usergroupid=10 or usergroupid=1 or usergroupid=4  or usergroupid=5  or usergroupid=6  or usergroupid=7  or usergroupid=8  or usergroupid=9  or usergroupid=24  or usergroupid=26) then
		if  rsnews("viewd")<>"" then  
		 V= Split(rsnews("viewd"),",") 
			For I = 0 To Ubound(V) 
			   if cint(V(I))=cint(session("userid")) then 
				   isviewd=true
				   'info=info&cint(V(I))&"-----"&Ubound(V)&"-"&I
				   exit for    
			   end if 	   
			Next 
			if isviewd=false then 
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
			end if 
		else
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
		end if 
	end if 
	
	
	if rsnews("viewdgroup")=3 then
		if  rsnews("viewd")<>"" then  
		 V= Split(rsnews("viewd"),",") 
			For I = 0 To Ubound(V) 
			   if cint(V(I))=cint(session("userid")) then 
				   isviewd=true
				   'info=info&cint(V(I))&"-----"&Ubound(V)&"-"&I
				   exit for    
			   end if 	   
			Next 
			if isviewd=false then 
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
			end if 
		else
				numb=numb+1
				info=info&"<LI><a href=news_view.asp?id="&rsnews("id")&" target=_blank>"&rsnews("news_title")&"</A></LI>"
		end if 
	end if 
	
	
		
	rsnews.movenext
	loop
	if numb<>0 then 
	   info="<span class='red'>未学习内容(共"&numb&"条)：</span></br>"&info
	   isinfo=true
	end if    
end if
rsnews.close
set rsnews=nothing



if isinfo then



info="<marquee  scrollamount=2 height=100 onmouseover=stop()  onmouseout=start() direction='up'>"&info&"</marquee>"
%>
<div id="newnotice">
	<p>
		<span class="title">提示信息</span>
        <span id="bts">
            <label class="button" id="tomin" title="最小化"> </label>
            <label class="button" id="tomax" title="最大化"> </label>
            <label class="button" id="toclose" title="关闭"> </label>
        </span>
    </p>
	<div id="noticecon"><%=info%></div>
</div>
<%end if%>  

</div>
</body>
</html>