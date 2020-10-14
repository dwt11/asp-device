<%@language=vbscript codepage=936 %>
<%
Option Explicit
response.buffer = True
Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<html>
<head>
<title>系统管理首页</title>
<style type="text/css">
<!--
.style1 {
	padding-TOP: 30px;
	padding-left:60px;

}
-->
</style>
<link rel="stylesheet" type="text/css" href="css/right.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body class='style1'>



<div class="col">
 
 
  <div class="block">
    <div align="center"><h3 class="block-title">欢迎登录 信息管理系统</h3></div>
<!--    <div class="block-body">请点击左侧菜单使用相关功能<br />
   </div>--> 
  </div>
  
  
  
  
  <div class="block">
    <div align="center"><h3 class="block-title">快捷方式</h3></div>
    <div class="block-body">
      <ul class="list">
        <%
		'dim leftmdb,connleft,connl
dim rs,sql,leftnumb
'leftmdb="ybdata/left.mdb"
Set connleft = Server.CreateObject("ADODB.Connection")
'connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
connleft.Open connl    

dim sqllog,rslog
sqllog="SELECT  * from left_class where zclass=0 ORDER BY orderby aSC"
set rslog=server.createobject("adodb.recordset")
rslog.open sqllog,connleft,1,1
if rslog.eof and rslog.bof then 
dwt.out "<p align='center'>没有任何日志</p>" 
else
           do while not rslog.eof 
                
                 if displaypagelevelh(session("groupid"),0,rslog("id")) then 
					  
dwt.out "<li><b><a href='left.asp?action=lefturlchick&pagelevelid="&rslog("id")&"&url="&rslog("url")&"' target=main>"&rslog("name")&"</a>:</b>"
		if rslog("id")=125 then dwt.out " <A href=dgtzl.asp?a=0>党委</a> "				
					dim sqlz,rsz,pagelevelid
					sqlz="SELECT  * from left_class where zclass="&rslog("id")&" ORDER BY orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,connleft,1,1
					if rsz.eof and rslog.bof then 
					dwt.out "<p align='center'>没有任何日志</p>" 
					else
				
   do while not rsz.eof 
						   if rslog("isbiglevel") then 
						     pagelevelid=rslog("id")
						   else
						     pagelevelid=rsz("id")
						   end if
	 
						   if displaypagelevelh(session("groupid"),0,pagelevelid) and rsz("isput") then 
							   dwt.out "&nbsp;&nbsp;&nbsp;&nbsp;<a href='left.asp?action=lefturlchick&pagelevelid="&pagelevelid&"&url="&rsz("url")&"' target=main>"
							   if rsz("isshartcut") then 
								 dwt.out "<font color='#ff0000'>"&rsz("name")&"</font>"
								else
									dwt.out rsz("name")
                                                                        
								end if  
							   dwt.out"</a>"

                                                          
						   end if 
						  
						rsz.movenext
						loop
				   end if
					   rsz.close
					   set rsz=nothing
					
					dwt.out " </li>"
               end if  
          rslog.movenext
          loop
   end if
       rslog.close
       set rslog=nothing

connleft.close
set connleft=nothing%>
      </ul>
    </div>
  </div>
  
  
  
  
    <%dwt.out "<div class='block'>"%>
    <h3 class="block-title">最近七天更新的内容</h3>
    <div class="block-body">
      <ul class="list">
        <%

sqllog="SELECT * from fdbw where now()-update<7"
set rslog=server.createobject("adodb.recordset")
rslog.open sqllog,connjg,1,1
if rslog.eof and rslog.bof then 
  dwt.out "" 
else
  dwt.out "<li><a href=left.asp?action=lefturlchick&pagelevelid=14&url=fdbw.asp?update=update> 防冻保温 </a></li>"           
end if
rslog.close
set rslog=nothing

sqllog="SELECT * from lsda where now()-update<7"
set rslog=server.createobject("adodb.recordset")
rslog.open sqllog,connjg,1,1
if rslog.eof and rslog.bof then 
  dwt.out "" 
else
  dwt.out "<li><a href=left.asp?action=lefturlchick&pagelevelid=13&url=lsda.asp?update=update>联锁档案</a></li>"           
end if
rslog.close
set rslog=nothing


sqllog="SELECT * from sb where now()-sb_update<7"
set rslog=server.createobject("adodb.recordset")
rslog.open sqllog,conn,1,1
if rslog.eof and rslog.bof then 
  dwt.out "" 
else
  dwt.out "<li><a href=left.asp?action=lefturlchick&pagelevelid=4&url=sb.asp?update=update*sbclassid="&rslog("sb_dclass")&">设备管理</a></li>"           
end if
rslog.close
set rslog=nothing

%>
      </ul>
    </div>
  </div>

  
  
      <%dwt.out "<div class='block'>"%>
    <h3 class="block-title">需要学习的内容</h3>
    <div class="block-body">
      <ul class="list">
        <%
dim sqlnews,rsnews,i
i =0
sqlnews="SELECT top 10 * from xzgl_news where isviewd order by id desc"
set rsnews=server.createobject("adodb.recordset")
rsnews.open sqlnews,connxzgl,1,1
if rsnews.eof and rsnews.bof then 
  dwt.out "" 
else
  do while not rsnews.eof
    Dwt.out "<li><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&rsnews("news_title")&"</a>&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
	i=i+1
   rsnews.movenext
   loop
end if
rsnews.close
set rsnews=nothing
if i>7 then Dwt.out "<Div align=right><a href=study_d.asp>更多学习内容....</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Div>"
%>
      </ul>
    </div>
  </div>

  
  
  
  
  
  
  
</div>










</body>
</html>
