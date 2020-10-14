<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->



<%
'数据库中 txdate字段为用户所选值班时间，txdate1为实际添写的时间，默认生成
dim sqlzblog,rszblog,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel,m,mi,mj
classid=request("classid")
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统--工作日志</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/tab.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	

sub add()
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>添加工作日志</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='gzrz.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	

   
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>姓名:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&session("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>日志内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	   Dwt.out "<iframe src='neweditor/editor.htm?id=body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"

	
	
	 dwt.out "  <input type='hidden' name='body' value=''>"	

	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='wlnumb' type='hidden' value='"&i&"'><input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""location.href='gzrz.asp';"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	
	
end sub	

sub saveadd()    
	 
	  '保存
 '  if request("body")<>"" then 
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from gzrz" 
      rsadd.open sqladd,connxzgl,1,3
      rsadd.addnew
      rsadd("body")=Trim(request("body"))
      rsadd("userid")=session("userid")
      rsadd("txdate")=request("txdate")
	  dwt.savesl "工作日志","添加",request("txdate")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
 ' end if 
	 
	 
		
		
		
		

	  dwt.out "<Script Language=Javascript>location.href='gzrz.asp?year="&year(request("txdate"))&"&month="&month(request("txdate"))&"&day="&day(request("txdate"))&"';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from gzrz where id="&id
   rsedit.open sqledit,connxzgl,1,1
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>编辑工作日志</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='gzrz.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>姓名:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&rsedit("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("txdate")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>日志内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	scontent=rsedit("body")
		'DWT.OUT "<input type='hidden' name='news_body' id='news_body' value='"&scontent&"'>"& vbCrLf
	    Dwt.out "<iframe src='neweditor/editor.htm?id=body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
	
	
    'dwt.out "<iframe src='neweditor/editor.htm?id=news_body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"


   dwt.out "<textarea name='body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf







	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'> <input name='id' type='hidden' value='"&request("id")&"'>    <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""location.href='gzrz.asp';"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	
	


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from gzrz where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connxzgl,1,3
rsedit("body")=Trim(request("body"))
'rsedit("userid")=session("userid")
rsedit("txdate")=request("txdate")
rsedit.update
rsedit.close
      
	  	dwt.savesl "工作日志","编辑",request("txdate")
	  dwt.out "<Script Language=Javascript>location.href='gzrz.asp?year="&year(request("txdate"))&"&month="&month(request("txdate"))&"&day="&day(request("txdate"))&"';</Script>"
	
end sub


sub main()
	url=geturl
	useridblog=request("userid")
	getyear=request("year")
	getmonth=request("month")
	getday=request("day")

    getnowday=date()-1
	
	if getyear="" then getyear=year(getnowday)
	if getmonth="" then getmonth=month(getnowday)
	if getday="" then getday=day(getnowday)



	selectdate=getyear&"-"&getmonth&"-"&getday
	selectdate=cdate(selectdate)
	'message selectdate
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><b>"&selectdate&" 工作日志</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
	dwt.out "		 <form method='post'  action='gzrz.asp'  name='form' >"
	'if session("level")=3 then 
    	dwt.out "		 <a href='/gzrz.asp?action=add'>添加日志</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	'end if 
	dwt.out "<a href='/gzrz.asp?year="&year(selectdate-2)&"&month="&month(selectdate-2)&"&day="&day(selectdate-2)&"'>"&year(selectdate-2)&"年"&month(selectdate-2)&"月"&day(selectdate-2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<a href='/gzrz.asp?year="&year(selectdate-1)&"&month="&month(selectdate-1)&"&day="&day(selectdate-1)&"'>"&year(selectdate-1)&"年"&month(selectdate-1)&"月"&day(selectdate-1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	dwt.out "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >		 <select name='year'></select>年<select name='month'></select>月<select name='day'></select>日 &nbsp;&nbsp;<input  type='submit' name='Submit' value=' 查看 ' style='cursor:hand;'>"
	dwt.out "		 <script type='text/javascript' src='js/selectdate.js'></script>"
	if now()-selectdate>1 then 	dwt.out "<a href='/gzrz.asp?year="&year(selectdate+1)&"&month="&month(selectdate+1)&"&day="&day(selectdate+1)&"'>"&year(selectdate-1)&"年"&month(selectdate+1)&"月"&day(selectdate+1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-selectdate>2 then 	dwt.out "<a href='/gzrz.asp?year="&year(selectdate+2)&"&month="&month(selectdate+2)&"&day="&day(selectdate+2)&"'>"&year(selectdate+2)&"年"&month(selectdate+2)&"月"&day(selectdate+2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	
	dwt.out "	</form></div>"
	dwt.out "</div>"

   
   
   
	

	dwt.out "<div class='navg'>"
	
	
	sqlstr="SELECT distinct userid from gzrz where  year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" order by userid"
	set rsstr=server.createobject("adodb.recordset")
	rsstr.open sqlstr,connxzgl,1,1
	if rsstr.eof and rsstr.bof then 
		dwt.out  message ("<p align='center'>未添加记录</p>" )
	else
		mi=0
	do while not rsstr.eof 
	dwt.out "  <div id='system' class='mainNavg'>"
		dwt.out "    <ul>"

	for mj=mi*8+1 to (mi+1)*8
		useridblog1= rsstr("userid") 
		
		
		  if useridblog="" then useridblog=useridblog1
		
		 
		  
		  
		  
		  if cint(useridblog)=rsstr("userid") then 
		   dwt.out "<li id='systemNavg'><a href='#'>"&usernameh(rsstr("userid"))&"</a></li>"
		  else
		   
		   dwt.out "<li><a href='gzrz.asp?userid="&rsstr("userid")&"&year="&getyear&"&month="&getmonth&"&day="&getday&"'>"&usernameh(rsstr("userid"))&"</a></li>"
		  end if 
          rsstr.movenext
		  if rsstr.eof then
		  exit for
		  end if
		  next
    dwt.out "</ul>"
    dwt.out " </div>"
		dwt.out "  <div class='textbody1'>"
    dwt.out " </div>"
		  mi=mi+1
		  if rsstr.eof then
		  exit do
		  end if
	loop
	end if 
	
	
	
	'if useridblog="" then useridblog=useridblog1
	
	
	
    
	  
	  
	
	dwt.out "  <div class='textbody' style='padding-top:0px;'>"
	
	  if useridblog="" then useridblog=0



                '上午
				sqlzblog="SELECT * from gzrz where userid="&useridblog&" and year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" ORDER BY id aSC "
				set rszblog=server.createobject("adodb.recordset")
				rszblog.open sqlzblog,connxzgl,1,1
				if rszblog.eof and rszblog.bof then 
                    'dwt.out "<span style='font-size:14px;color:#0000ff;font-weight: bold;'>"&rsssbz("bzname")&"</span><br><br>"
					dwt.out  "<div class='textbody1'>未添写 "&selectdate&" 日志</div>"
				else
					'dwt.out "<span style='font-size:14px;color:#0000ff;font-weight: bold;'>"&rsssbz("bzname")&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
					'dwt.out "<b>"&selectdate&" 白班</b> 值班人:<b>"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)&"</b> 更新时间:"&rszblog("txdate1")
					dwt.out " <br>"
				do while not rszblog.eof 
					dwt.out "<div class='textbody1'>"&DecodeFilter(rszblog("body"),"FONT,STRONG")&"&nbsp;&nbsp;&nbsp;&nbsp;"
					 if session("userid")=325 or rszblog("userid")=session("userid") then 
						 dwt.out  "<a href='gzrz.asp?action=edit&id="&rszblog("id")&"'>编辑</a>&nbsp;"
						 dwt.out "<a href='gzrz.asp?action=del&id="&rszblog("id")&"' onClick=""return confirm('确定要删除此内容吗？');"">删除</a>"
					 end if 
					
					
					'call editdel(rszblog("id"),rszblog("sscj"),rszblog("userid"),"zblog.asp?action=edit&id=","zblog.asp?action=del&id=")
				    dwt.out "</div>"
				rszblog.movenext
				loop
				end if 
				rszblog.close	
           
		   
		   
		   
		       

	dwt.out "</div>"
	dwt.out "</div>	"
end sub	

	
	
	
	
	
	
	
	



sub del()
ID=request("ID")
	sqledit="select * from gzrz where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connxzgl,1,1
	
	dwt.savesl "工作日志","删除",rsedit("txdate")
	rsedit.close



set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from gzrz where id="&id
rsdel.open sqldel,connxzgl,1,3
dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub





dwt.out  "</body></html>"

Call CloseConn
%>