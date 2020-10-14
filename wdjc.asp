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
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统--发热设备巡检记录</title>"& vbCrLf
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
  case "del"
    'if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
	call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then 
	call main
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
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>添加巡检记录</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='wdjc.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

    if session("level")=3 then 
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所班组:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&ssbzh(session("levelzclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='ssbz' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>添加人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&session("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 75px'>巡检表:</LABEL>"& vbCrLf
		dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	  sqlclass="SELECT * from address where ssbz="&session("levelzclass")&" order by id"
	  set rsclass=server.createobject("adodb.recordset")
	  rsclass.open sqlclass,connw,1,1
	  if rsclass.eof and rsclass.bof then 
		  dwt.out  message ("<p align='center'>未添加巡检设备</p>" )
	  else
	  
	  
	  
		  dwt.out "<table   border='1'  cellpadding='1' cellspacing='1'>"
	  	  dwt.out "<tr>"
		  
		  dwt.out "<td align=center>编号</td>"
	  	  dwt.out "<td align=center>位置</td>"
	  	  dwt.out "<td align=center>名称</td>"
	  	  dwt.out "<td align=center>温度</td>"
		  dwt.out "</tr>"
	  do while not rsclass.eof 
	  
	  
	  	  dwt.out "<tr>"
		  
		  dwt.out "<td>"&rsclass("id")&"</td><input name='bh"&rsclass("id")&"' type='hidden' value='"&rsclass("id")&"'>"
		  
	  	  dwt.out "<td>"&rsclass("wz")&"</td><input name='wz"&rsclass("id")&"' type='hidden' value='"&rsclass("wz")&"'>"
	  	  dwt.out "<td>"&rsclass("name")&"</td><input name='name"&rsclass("id")&"' type='hidden' value='"&rsclass("name")&"'>"
	  	  dwt.out "<td><input name='ti"&rsclass("id")&"' type='text' ></td>"
		  dwt.out "</tr>"
	  
	  rsclass.movenext
	  loop
	  dwt.out "</table>"
	  end if 

		
		
		
		
		
		
		
		
		
		
		
dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""location.href='';"" style='cursor:hand;'>"& vbCrLf
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
      
	  
	  
	  
	  sqlclass="SELECT * from address"
	  set rsclass=server.createobject("adodb.recordset")
	  rsclass.open sqlclass,connw,1,1
	  if rsclass.eof and rsclass.bof then 
		  'dwt.out  message ("<p align='center'>未添加巡检设备</p>" )
	  else
	  do while not rsclass.eof 
		  
		  if request("wz"&rsclass("id"))<>"" and  request("name"&rsclass("id"))<>"" and request("ti"&rsclass("id"))<>"" then
		  
			
			set rsadd=server.createobject("adodb.recordset")
			sqladd="select * from bb" 
			rsadd.open sqladd,connw,1,3
			rsadd.addnew
			rsadd("userid")=session("userid")
			rsadd("ssbz")=session("levelzclass")
			rsadd("wz")=request("wz"&rsclass("id"))
			rsadd("name")=request("name"&rsclass("id"))
			rsadd("ti")=request("ti"&rsclass("id"))
			rsadd("update")=now()
			rsadd.update
			rsadd.close
			set rsadd=nothing
		  end if 
	  rsclass.movenext
	  loop
	  end if
	  
	  
	  dwt.savesl "温度巡检记录","添加",now()
	 
	 
		
		
		

	 dwt.out "<Script Language=Javascript>location.href='?ssbz="&session("levelzclass")&"';</Script>"
end sub



sub main()
	url=GetUrl
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><b>发热设备检查记录</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
    	dwt.out "		 <a href='?action=add'>添加记录</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='wdjc_class.asp'>巡检设备管理</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='wdjc_class.asp?action=add'>添加巡检设备</a>"
'	
	
	dwt.out "	使用说明：以包机人用户名登录后，点击“添加巡检设备”，添加完成后点击“添加记录”在对应的设备里添写温度即可。巡检设备只需要添加一次即可</div>"
	dwt.out "</div>"

   
   
   
	

	dwt.out "<div class='navg'>"
	dwt.out "  <div id='system' class='mainNavg'>"
	dwt.out "    <ul>"
		sscjid=request("sscj")
		'101218修改，打开页面后自动显示对应的车间
		 
		 
		  if sscjid="" then 
		     if  request("ssbz")="" then 
    			  sscjid=1    '101218修改，打开页面后自动显示对应的车间
			 else
					sqlbz="SELECT * from bzname where id="&request("ssbz")
					set rsbz=server.createobject("adodb.recordset")
					rsbz.open sqlbz,conn,1,1
					if rsbz.eof and rsbz.bof then 
						'dwt.out  message ("<p align='center'>未添加班组</p>" )
sscjid=0
					else
						sscjid=rsbz("sscj")
					end if 

			 end if 	  
			  
		  end if 	  
			  
	
	
	
	sqlsscj="SELECT * from levelname where levelclass=1 and levelid<4"
	set rssscj=server.createobject("adodb.recordset")
	rssscj.open sqlsscj,conn,1,1
	if rssscj.eof and rssscj.bof then 
		dwt.out  message ("<p align='center'>未添加生产车间</p>" )
	else
	do while not rssscj.eof 
		if cint(sscjid)=rssscj("levelid") then 
		   dwt.out "<li id='systemNavg'><a href='#'>"&rssscj("levelname")&"</a></li>"
		else
		   dwt.out "<li><a href='?sscj="&rssscj("levelid")&"'>"&rssscj("levelname")&"</a></li>"
		end if    
	rssscj.movenext
	loop
	end if 
	  
	  
    dwt.out "</ul>"
    dwt.out " </div>"
	
	dwt.out "  <div class='textbody'>"
		sqlssbz="SELECT * from bzname where sscj="&sscjid
		set rsssbz=server.createobject("adodb.recordset")
		rsssbz.open sqlssbz,conn,1,1
		if rsssbz.eof and rsssbz.bof then 
			'dwt.out  message ("<p align='center'>未添加班组</p>" )
		else
		
		
		
		
		do while not rsssbz.eof 



	  dwt.out "<b><a href='?ssbz="&rsssbz("id")&"'>"&rsssbz("bzname")&"</a></b> "
	  ij=ij+1
	  if ij=1 then 
		ssbzid=rsssbz("id")
		ssbzname=rsssbz("bzname")
	  end if 
	  
	rsssbz.movenext
	loop
	end if 
	
	  if request("ssbz")<>"" then 
		ssbzid=request("ssbz")
			  sqlbz="SELECT * from bzname where id="&ssbzid
			  set rsbz=server.createobject("adodb.recordset")
			  rsbz.open sqlbz,conn,1,1
			  if rsbz.eof and rsbz.bof then 
				 ' dwt.out  message ("<p align='center'>未添加班组</p>" )
			  else
				  ssbzname=rsbz("bzname")
			  end if 
	  end if 
	
	
	


		  dwt.out "<br><table  border='1'  cellpadding='1' cellspacing='1'>"
		  dwt.out "		  <tr>"
		  dwt.out "		    <td colspan='5' align=center><b>"&ssbzname&"</b></td>"
		  dwt.out "	      </tr>"
			  dwt.out "		  <tr>"
			  dwt.out "		    <td align=center>日期</td>"
			  dwt.out "		    <td align=center>位置</td>"
			  dwt.out "		    <td align=center>名称</td>"
			  dwt.out "		    <td align=center>温度</td>"
			  dwt.out "		    <td>&nbsp;</td>"
			  dwt.out "	      </tr>"



		  sqljl="SELECT * from bb where ssbz="&ssbzid&" order by wz,update desc" 
		  set rsjl=server.createobject("adodb.recordset")
		  rsjl.open sqljl,connw,1,1
		  if rsjl.eof and rsjl.bof then 
		  dwt.out "		  <tr>"
		  dwt.out "		    <td colspan='5' align=center>未添加记录</td>"
		  dwt.out "	      </tr>"
		  else
		  record=rsjl.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
	   end if 
	   rsjl.PageSize = Cint(PgSz) 
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
	   rsjl.absolutePage = page
	   start=PgSz*Page-PgSz+1
	   rowCount = rsjl.PageSize
	   do while not rsjl.eof and rowcount>0
			 ' dwt.out rsjl("update")&" "&rsjl("wz")&" "&rsjl("ti")&"<br>"
			  dwt.out "		  <tr>"
			  dwt.out "		    <td>"&rsjl("update")&"</td>"
			  dwt.out "		    <td>"&rsjl("wz")&"</td>"
			  dwt.out "		    <td>"&rsjl("name")&"</td>"
			  dwt.out "		    <td>"&rsjl("ti")&"</td>"
			  dwt.out "		    <td><a href=wdjc_view.asp?name="&Server.URLEncode(rsjl("name"))&"&ssbz="&rsjl("ssbz")&"&wz="&Server.URLEncode(rsjl("wz"))&"  target='_blank'>看此位置所有记录</a>   "
				if session("level")=1 or session("level")=0 and session("levelclass")=sscj then  response.write "<a href=?action=del&id="&rsjl("id")&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
			
			  dwt.out "</td>"
			  dwt.out "	      </tr>"
		
		RowCount=RowCount-1
		  rsjl.movenext
		  loop
		  end if 
	
		  dwt.out "</table><br>"
	
if request("ssbz")<>"" or request("sscj")<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 	
	
	
	
	
	
	
	
	dwt.out "</div>"
	dwt.out "</div>	"
	
	
	rsjl.close
	set rsjl=nothing
	conn.close
	set conn=nothing
end sub	

	
	
	
	
	
	
	
	

sub del()
ID=request("ID")



set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from bb where id="&id
rsdel.open sqldel,connw,1,3
dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  
dwt.savesl "温度巡检记录","删除",now()

end sub






dwt.out  "</body></html>"

Call CloseConn
%>
