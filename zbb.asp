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
dim sqlzbb,rszbb,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
url="zbb.asp"

dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统值班表管理页</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf

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
   dwt.out "<br><form method='post' action='zbb.asp' name='form1'>"
   dwt.out "<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>添加值班表</strong></div></td>    </tr>"
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out "<strong>标&nbsp;&nbsp;题：</strong></td>"
	 dwt.out "<td width='88%' class='tdbg'><input name='zbb_title' type='text'></td>    </tr>   "
			 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>时&nbsp;&nbsp;间：</strong></td> "
   dwt.out "<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out "<input name='zbb_date' type='text' value="&now()&" >"
   dwt.out "<a href='#' onClick=""popUpCalendar(this,pxst_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out "<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    dwt.out "<td width='88%' class='tdbg'>"
	 dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=zbb_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
      dwt.out "</iframe>  <input type='hidden' name='zbb_body' value=''>"
    dwt.out "</td></tr>  "   
    dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='zbb.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	

sub saveadd()    
	  '保存新增用户
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zbb" 
      rsadd.open sqladd,conna,1,3
      rsadd.addnew
      rsadd("title")=ReplaceBadChar(Trim(Request("zbb_title")))
      rsadd("body")=Trim(request("zbb_body"))
      rsadd("date")=request("zbb_date")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out "<Script Language=Javascript>location.href='zbb.asp';</Script>"
end sub

sub edit()
     '编辑用户
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zbb where id="&id
   rsedit.open sqledit,conna,1,1

   dwt.out "<br><form method='post' action='zbb.asp' name='form1'>"
   dwt.out "<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>编辑值班表</strong></div></td>    </tr>"
   dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
   dwt.out "<strong>标&nbsp;&nbsp;题：</strong></td>"
   dwt.out "<td width='88%' class='tdbg'><input name='zbb_title' type='text' value='"&rsedit("title")&"'></td>    </tr>   "
		 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>时&nbsp;&nbsp;间：</strong></td> "
   dwt.out "<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out "<input name='zbb_date' type='text' value="&rsedit("date")&" >"
   dwt.out "<a href='#' onClick=""popUpCalendar(this,pxst_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out "<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
  
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>内&nbsp;&nbsp;容： </strong></td>"      
    dwt.out "<td width='88%' class='tdbg'>"
	scontent=rsedit("body")
	 dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=zbb_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      dwt.out "</iframe><textarea name='zbb_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
    dwt.out "</td></tr>  "   

	 
    dwt.out "<tr> <td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='zbb.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from zbb where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conna,1,3
rsedit("title")=ReplaceBadChar(Trim(Request("zbb_title")))
rsedit("body")=Trim(request("zbb_body"))
      rsedit("date")=request("zbb_date")
rsedit.update
rsedit.close
	  dwt.out "<Script Language=Javascript>history.go(-2);</Script>"
	
end sub


sub main()
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  " <tr class='topbg'>"& vbCrLf
dwt.out  "   <td height='22' colspan='2' align='center'><strong>值班表管理页</strong></td>"& vbCrLf
dwt.out  "  </tr>  "& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
dwt.out  "    <td height='30'><a href='zbb.asp'>值班表管理首页</a>"
if session("levelclass")=10 or session("levelclass")=9 then dwt.out  "&nbsp;|&nbsp;<a href='zbb.asp?action=add'>添加值班表</a>"
dwt.out  "    </td>"& vbCrLf
dwt.out  "  </tr>"& vbCrLf
dwt.out  "</table>"& vbCrLf

dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out  "<tr class=""title"">" 
dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><div align=""center""><strong>值班表标题</strong></div></td>"
dwt.out  "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选项</strong></div></td>"
dwt.out  "    </tr>"
sqlzbb="SELECT * from zbb ORDER BY id DESC"
set rszbb=server.createobject("adodb.recordset")
rszbb.open sqlzbb,conna,1,1
if rszbb.eof and rszbb.bof then 
dwt.out  "<p align='center'>未添加值班表</p>" 
else
           record=rszbb.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rszbb.PageSize = Cint(PgSz) 
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
           rszbb.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rszbb.PageSize
           do while not rszbb.eof and rowcount>0
                 xh=xh+1
                 dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                 dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
                 dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""40%""><a href=zbb_view.asp?id="&rszbb("id")&" target=_blank>"&rszbb("title")&"</a></td>"
                 dwt.out  "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 if session("level")=0 or session("level")=9 or session("level")=2 then dwt.out  "<a href='zbb.asp?action=edit&ID="&rszbb("id")&"'>编辑</a>&nbsp;  <a href='zbb.asp?action=del&ID="&rszbb("id")&"' onClick=""return confirm('确定要删除此内容吗？');"">删除</a>"
				 dwt.out  "&nbsp;</div></td>"
                 dwt.out  "    </tr>"
                 RowCount=RowCount-1
          rszbb.movenext
          loop
       end if
       rszbb.close
       set rszbb=nothing
        conn.close
        set conn=nothing
        dwt.out  "</table>"
       call showpage1(page,url,total,record,PgSz)
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from zbb where id="&id
rsdel.open sqldel,conna,1,3
dwt.out "<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub







dwt.out  "</body></html>"

Call CloseConn
%>