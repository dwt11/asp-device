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


<%
dim record,pgsz,total,page,start,rowcount,ii
dim url,xh

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统-主要设备运转表</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

if request("action")="" then call main 
if request("action")="zysbname" then call zysbname
if request("action")="addsb" then call addsb
if request("action")="saveaddsb" then call saveaddsb
if request("action")="editsb" then call editsb
if request("action")="saveeditsb" then call saveeditsb
if request("action")="delsb" then call delsb



'*****************************************************
'列表每个月各车间报的运转率
'**********************************************88888888888
sub main()    
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>主要设备运转表</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""zysbyz.asp"">主要设备运转率首页</a>&nbsp;|&nbsp;<a href=""zysbyz_view.asp?action=add"">添加运转率</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=zysbname"">设备管理</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=addsb"">添加设备</a>"& vbCrLf
response.write " </td> </tr>"& vbCrLf
response.write "</table>"& vbCrLf

url="zysbyz.asp"
dim sqlzysbyz,rszysbyz
sqlzysbyz="SELECT  distinct sscj,year,month  from zysbyz  ORDER BY month DESC"
set rszysbyz=server.createobject("adodb.recordset")
rszysbyz.open sqlzysbyz,connb,1,1
if rszysbyz.eof and rszysbyz.bof then 
message("未添加内容")
else

response.write "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>日期</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
           record=rszysbyz.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rszysbyz.PageSize = Cint(PgSz) 
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
           rszysbyz.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rszysbyz.PageSize
           do while not rszysbyz.eof and rowcount>0
		xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center"">"&xh&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><a href=zysbyz_view.asp?month="&rszysbyz("month")&"&sscj="&rszysbyz("sscj")&"&year="&rszysbyz("year")&">"&sscjh(rszysbyz("sscj"))&"</a></div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center"">"&rszysbyz("year")&"年"&rszysbyz("month")&"月</div></td>"
                response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""></div></td></tr>"
                 RowCount=RowCount-1
          rszysbyz.movenext
          loop
        response.write "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rszysbyz.close
       set rszysbyz=nothing
        connb.close
        set connb=nothing
end sub











'*****************************************************
'主要设备名称位号的管理,显示\添加\编辑\删除
'**********************************************88888888888



sub addsb()
dim rscj,sqlcj
   response.write"<br><br><br><form method='post' action='zysbyz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>添加主要设备</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
   response.write"<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	response.write"<select name='zysbname_sscj' size='1'>"
    response.write"<option >选择所属车间</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select></td></tr>  "  	 
  else 	 
    response.write"<input name='zysbname_sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    response.write"<input name='zysbname_sscj' type='hidden' value="&session("level")&"></td></tr>"& vbCrLf
 end if 
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
   response.write"<strong>位&nbsp;&nbsp;号：</strong></td>"
   response.write"<td width='88%' class='tdbg'><input name='zysbname_wh' type='text'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>名&nbsp;&nbsp;称：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='zysbname_name' ></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数&nbsp;&nbsp;量：</strong></td> "
	response.write"<td><select name='zysbname_numb' size='1'>"
	dim i
	for i=1 to 20
		response.write"<option value='"&i&"'>"&i&"</option>"
    next
	response.write"</select></td></tr>"
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveaddsb'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveaddsb()    
	dim rsadd,sqladd
	'  on error resume next
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zysbname" 
      rsadd.open sqladd,connb,1,3
      rsadd.addnew
	  
      rsadd("sscj")=Trim(Request("zysbname_sscj"))
      rsadd("wh")=request("zysbname_wh")
      rsadd("name")=request("zysbname_name")
      rsadd("numb")=request("zysbname_numb")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='zysbyz.asp?action=zysbname';</Script>"
end sub

sub delsb()
dim id,rsdel,sqldel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from zysbname where id="&id
  rsdel.open sqldel,connb,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub

sub saveeditsb()   

dim sqledit,rsedit 
      'on error resume next
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zysbname where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connb,1,3
      rsedit("sscj")=Trim(Request("zysbname_sscj"))
      rsedit("wh")=request("zysbname_wh")
	  rsedit("name")=request("zysbname_name")
	  rsedit("numb")=request("zysbname_numb")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub editsb()
  dim id,sqledit,rsedit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zysbname where id="&id
   rsedit.open sqledit,connb,1,1
   response.write"<br><br><br><form method='post' action='zysbyz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>编辑主要设备</strong></div></td>    </tr>"
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     response.write"<td width='88%' class='tdbg'><input name='zysbname_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     response.write"<input name='zysbname_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>位&nbsp;&nbsp;号：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='zysbname_wh' type='text' value='"&rsedit("wh")&"'></td>    </tr>   "
	 	 
		 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>名&nbsp;&nbsp;称：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='zysbname_name' type='text' value='"&rsedit("name")&"'></td>    </tr>   "


 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数&nbsp;&nbsp;量：</strong></td> "
	response.write"<td><select name='zysbname_numb' size='1'>"
	dim i
	for i=1 to 20
		response.write"<option value='"&i&"' "
		if rsedit("numb")=i then response.write"selected"
		response.write">"&i&"</option>"
    
	next
	
	   
	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveeditsb'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
	       rsedit.close
       set rsedit=nothing
	

end sub

sub zysbname()
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>主要设备运转表</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""zysbyz.asp"">主要设备运转率首页</a>&nbsp;|&nbsp;<a href=""zysbyz_view.asp?action=add"">添加运转率</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=zysbname"">设备管理</a>&nbsp;|&nbsp;<a href=""zysbyz.asp?action=addsb"">添加设备</a>"& vbCrLf
response.write " </td> </tr>"& vbCrLf
response.write "</table>"& vbCrLf

url="zysbyz.asp?action=zysbname"

   dim sqlzysbname,rszysbname
   sqlzysbname="SELECT * from zysbname  ORDER BY id DESC"
   set rszysbname=server.createobject("adodb.recordset")
   rszysbname.open sqlzysbname,connb,1,1
   if rszysbname.eof and rszysbname.bof then 
      message("未添加内容")
   else
      response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      response.write "<tr class=""title"">" 
      response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
      response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center""><strong>单&nbsp;&nbsp;位</strong></div></td>"
      response.write "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>位&nbsp;&nbsp;号</strong></div></td>"
      response.write "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>名&nbsp;&nbsp;称</strong></div></td>"
      response.write "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>数&nbsp;&nbsp;量</strong></div></td>"
      response.write "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选&nbsp;&nbsp;项</strong></div></td>"
      response.write "    </tr>"
      record=rszysbname.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rszysbname.PageSize = Cint(PgSz) 
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
           rszysbname.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rszysbname.PageSize
		   do while not rszysbname.eof and rowcount>0
		xh=xh+1
                response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
				response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center"">"&sscjh(rszysbname("sscj"))&"</div></td>"
				response.write "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px"">"&rszysbname("wh")&"</td>"
 				response.write "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px"">"&rszysbname("name")&"</td>"
				response.write "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszysbname("numb")&"</div></td>"
				response.write "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				call editdel(rszysbname("id"),rszysbname("sscj"),"zysbyz.asp?action=editsb&id=","zysbyz.asp?action=delsb&id=")
				response.write "</div></td></tr>"
          RowCount=RowCount-1
		   rszysbname.movenext
     loop
     
	 response.write "</table>"
	        call showpage(page,url,total,record,PgSz)

 end if
       rszysbname.close
       set rszysbname=nothing

end sub 




response.write "</body></html>"



Call CloseConn
%>