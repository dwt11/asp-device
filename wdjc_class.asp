<%@language=vbscript codepage=936 %>
<%
Option Explicit
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


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh
dim rsadd,sqladd,TrueIP,id,rsedit,sqledit,rsdel,sqldel
dim sqluser,rsuser,sqlcj,rscj
url="wdjc_class.asp"
action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then  call add
  case "saveadd"
    call saveadd
  case "edit"
	'if truepagelevelh(session("groupid"),2,session("pagelevelid")) then 
	call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then 
	call main
end select	
Dwt.out "<html>"
Dwt.out "<head>"
Dwt.out "<title>发热设备管理</title>"
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function CheckAdd(){" & vbCrLf
 Dwt.out " if(document.form1.class_name.value==''){" & vbCrLf
Dwt.out "      alert('名称不能为空！');" & vbCrLf
Dwt.out "   document.form1.class_name.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out "</head>"
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"


sub add()
   '新增
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>巡检设备添加</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='wdjc_class.asp' name='form1' >"
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
	dwt.out"				<LABEL style='WIDTH: 75px'>位置:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=wz>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>名称:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='name' class='x-form-text x-form-field' style='WIDTH: 175px' >"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value='  保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""location.href='';"" style='cursor:hand;'>"& vbCrLf
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
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from address" 
      rsadd.open sqladd,connw,1,3
      rsadd.addnew
rsadd("name")=ReplaceBadChar(Trim(Request("name")))
rsadd("wz")=ReplaceBadChar(Trim(Request("wz")))
			rsadd("ssbz")=session("levelzclass")
	 
	 
      rsadd.update
      rsadd.close
      set rsadd=nothing
	 dwt.out "<Script Language=Javascript>location.href='wdjc_class.asp';</Script>"
	
	  
end sub

sub main()
   	
 
	dim sqlsscj,rssscj,sscjd,sql,sqlssbz,rsssbz,idd,sscjdd
		sqlsscj="SELECT * from levelname where levelclass=1 and levelid<4"
		set rssscj=server.createobject("adodb.recordset")
		rssscj.open sqlsscj,conn,1,1
		if rssscj.eof and rssscj.bof then 
			'dwt.out  message ("<p align='center'>未添加生产车间</p>" )
		else
		do while not rssscj.eof 
			   sscjd=sscjd& "<a href='?sscj="&rssscj("levelid")&"'>"&rssscj("levelname")&"</a> "
		rssscj.movenext
		loop
		end if 
	
	if request("sscj")<>"" then 
				

			sqlssbz="SELECT * from bzname where sscj="&request("sscj")
			set rsssbz=server.createobject("adodb.recordset")
			rsssbz.open sqlssbz,conn,1,1
			if rsssbz.eof and rsssbz.bof then 
				'dwt.out  message ("<p align='center'>未添加班组</p>" )
			else
			
			
			
			
			do while not rsssbz.eof 
	        idd=idd+1
			if idd=1 then 
			  sql=" ssbz="&rsssbz("id")
			else
			sql=sql&" or  ssbz="&rsssbz("id")
   sscjdd=sscjh(request("sscj"))&"-"
	end if 
		  
		rsssbz.movenext
		loop
		end if 
	
	end if 	
	 	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&sscjdd&"发热设备管理</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
'用户管理首页
	Dwt.out "<Div class='x-toolbar'><Div align=left><a href='?action=add'>添加设备</a> "&sscjd&"</Div></Div>" & vbCrLf
 		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf

      if sql<>"" then  sqlbody="SELECT * from address where "&sql&" order by id desc"
      if sql="" then  sqlbody="SELECT * from address  order by id desc"
		
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,connw,1,1
      if rsbody.eof and rsbody.bof then 
           Dwt.out "<p align=""center"">暂无内容</p>" 
      else
	  Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      Dwt.out "<tr  class=""x-grid-header"">" 
      Dwt.out "     <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center""><strong>序号</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>车间</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>班组</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>位置</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>名称</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>操作</strong></Div></td>"
      Dwt.out "    </tr>"
           record=rsbody.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsbody.PageSize = Cint(PgSz) 
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
           rsbody.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsbody.PageSize
           do while not rsbody.eof and rowcount>0
              
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center"">"&xh_id&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&bzsscj(rsbody("ssbz"))&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&ssbzh(rsbody("ssbz"))&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("wz")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("name")&"</Div></td>"
                  
		if session("levelzclass")=rsbody("ssbz") then		
Dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><a href='?action=edit&ID="&rsbody("id")&"'>编辑</a>&nbsp;"
 Dwt.out "  <a href='?action=del&ID="&rsbody("id")&"' onClick=""return confirm('确定要删除吗？');"">删除</a></Div></td>"
else
dwt.out "<td></td>"
end if 
                 Dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
		Dwt.out "</table>"& vbCrLf
		call showpage1(page,url,total,record,PgSz)
		Dwt.out "</Div>"& vbCrLf
       end if
 	Dwt.out "</Div>"  
      rsbody.close
       set rsbody=nothing
       
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from address where id="&id
   rsedit.open sqledit,connw,1,1


  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>巡检设备编辑</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='wdjc_class.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所车间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&bzsscj(rsedit("ssbz"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

    if session("level")=3 then 
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >属所班组:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&ssbzh(rsedit("ssbz"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>位置:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=wz value='"&rsedit("wz")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>名称:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='name' class='x-form-text x-form-field' style='WIDTH: 175px'  value='"&rsedit("name")&"'>"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&rsedit("id")&"'>    <input  type='submit' name='Submit' value='  保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""location.href='';"" style='cursor:hand;'>"& vbCrLf
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
sqledit="select * from address where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connw,1,3
rsedit("name")=ReplaceBadChar(Trim(Request("name")))
rsedit("wz")=ReplaceBadChar(Trim(Request("wz")))
rsedit.update
rsedit.close
	
Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from address where id="&id
rsdel.open sqldel,connw,1,3
Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  
end sub

Dwt.out "</body></html>"



Function bzsscj(ssbz)
    dim sqlcj,rscj,sscj
	  sqlcj="SELECT * from bzname where id="&ssbz
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    if rscj.eof then 
       sscj=0
	else  
      sscj=rscj("sscj")
	end if

	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    if rscj.eof then 
	  bzsscj="未知"
	else  
	    bzsscj=rscj("levelname")
	end if
	rscj.close
	set rscj=nothing
end Function 

Call CloseConn
%>