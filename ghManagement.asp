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


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh,xh_id
dim rsadd,sqladd,TrueIP,id,rsedit,sqledit,rsdel,sqldel
dim sqluser,rsuser,sqlcj,rscj
url="ghmanagement.asp"
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
dwt.out "<html>"
dwt.out "<head>"
dwt.out "<title>装置管理</title>"
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function CheckAdd(){" & vbCrLf
 dwt.out " if(document.form1.ghname.value==''){" & vbCrLf
dwt.out "      alert('名不能为空！');" & vbCrLf
dwt.out "   document.form1.ghname.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "  if(document.form1.sscj.value==''){" & vbCrLf
dwt.out "      alert('未选择所属车间！');" & vbCrLf
 dwt.out "  document.form1.sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"



sub add()
   dwt.out"<form method='post' action='ghmanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>新 增 装 置</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>装 置 名：</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='ghname' type='text'></td>    </tr>   "
	dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属车间： </strong></td>"& vbCrLf      
    dwt.out"<td width='88%' class='tdbg'>"& vbCrLf
	
	dwt.out"<select name='sscj' size='1'>"& vbCrLf
    dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>"  	 & vbCrLf
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='usermanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from ghname" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
      rsadd("gh_name")=ReplaceBadChar(Trim(Request("ghname")))
      rsadd("sscj")=ReplaceBadChar(Trim(request("sscj")))
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub

sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>装置管理</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
    dwt.out "<div class='x-toolbar'>" & vbCrLf
    dwt.out "<div align=left><a href=ghmanagement.asp?action=add>添加装置</a></div>"
	dwt.out "</div>"
      sqlbody="SELECT * from ghname  ORDER BY SSCJ ASC,GH_NAME ASC"
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,conn,1,1
      if rsbody.eof and rsbody.bof then 
           dwt.out "<p align=""center"">暂无内容</p>" 
      else
		  	 dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
			dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			dwt.out "<tr  class=""x-grid-header"">" 
			dwt.out "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>装 置 名</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>所属车间</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>操作</div></td>"
			dwt.out "    </tr>"
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
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1

			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("gh_name")&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sscj"))&"</div></td>"
                  dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='ghmanagement.asp?action=edit&ID="&rsbody("ghid")&"'>编辑</a>&nbsp;"
				 dwt.out "  <a href='ghmanagement.asp?action=del&ID="&rsbody("ghid")&"' onClick=""return confirm('确定要删除此装置吗？删除后其他相关的此装置内容将不会显示');"">删除</a></div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
       end if
       rsbody.close
       set rsbody=nothing
        conn.close
        set conn=nothing
        dwt.out "</table>"
       call showpage1(page,url,total,record,PgSz)
	   dwt.out "</div></div>"
end sub

sub edit()
     '编辑用户
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from ghname where ghid="&id
   rsedit.open sqledit,conn,1,1

   dwt.out"<form method='post' action='ghmanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>编 辑 装 置</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>装置名：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'>"
	 dwt.out"<input type='text' name='ghname' value="&rsedit("gh_name")&"></td>    </tr> "
	dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属车间： </strong></td>"& vbCrLf      
    dwt.out"<td width='88%' class='tdbg'>"& vbCrLf
	
	dwt.out"<select name='sscj' size='1'>"& vbCrLf
    dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'"
		if rsedit("sscj")=rscj("levelid") then dwt.out"selected"
		dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>"  	 & vbCrLf
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from ghname where ghID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conn,1,3
rsedit("gh_name")=ReplaceBadChar(Trim(Request("ghname")))
rsedit("sscj")=ReplaceBadChar(Trim(Request("sscj")))
rsedit.update
rsedit.close
	
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from ghname where ghid="&id
rsdel.open sqldel,conn,1,3
dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub

dwt.out "</body></html>"

Call CloseConn
%>