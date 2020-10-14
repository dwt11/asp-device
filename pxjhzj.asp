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
dim sqlpxjh,rspxjh,sqlpxzj,rspxzj
dim record,pgsz,total,page,start,rowcount,xh,url,ii

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统培训管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if request("action")="pxjh" then call pxjh
if request("action")="pxzj" then call pxzj


sub pxjh()
response.write "<br><br><table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>培训计划页</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf


url="pxjhzj.asp?action=pxjh"
sqlpxjh="SELECT  distinct sscj,year,month  from pxjh order by year desc,month  DESC"
set rspxjh=server.createobject("adodb.recordset")
rspxjh.open sqlpxjh,conne,1,1
if rspxjh.eof and rspxjh.bof then 
message("未添加培训计划")
else

response.write "<table  width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>日期</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
           record=rspxjh.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxjh.PageSize = Cint(PgSz) 
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
           rspxjh.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxjh.PageSize
           do while not rspxjh.eof and rowcount>0
		xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center"">"&xh&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><a href=pxjh_view.asp?action=pxjh&month="&rspxjh("month")&"&sscj="&rspxjh("sscj")&"&year="&rspxjh("year")&">"&sscjh(rspxjh("sscj"))&"</a></div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center"">"&rspxjh("year")&"年"&rspxjh("month")&"月</div></td>"
                response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><a href=tocsv.asp?action=pxjhmain&titlename=培训计划&month="&rspxjh("month")&"&sscj="&rspxjh("sscj")&"&year="&rspxjh("year")&">导出到EXCEL文档</a></div></td></tr>"
                 RowCount=RowCount-1
          rspxjh.movenext
          loop
        response.write "</table>"
       call showpage_80(page,url,total,record,PgSz)
       end if
       rspxjh.close
       set rspxjh=nothing
        conne.close
        set conne=nothing



end sub

sub pxzj()
response.write "<br><br><table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>培训总结页</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf
url="pxjhzj.asp?action=pxzj"
sqlpxzj="SELECT  distinct sscj,year,month  from pxzj order by year desc, month  DESC"
set rspxzj=server.createobject("adodb.recordset")
rspxzj.open sqlpxzj,conne,1,1
if rspxzj.eof and rspxzj.bof then 
message("未添加培训总结")
else

response.write "<table  width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>日期</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
           record=rspxzj.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxzj.PageSize = Cint(PgSz) 
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
           rspxzj.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxzj.PageSize
           do while not rspxzj.eof and rowcount>0
		xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center"">"&xh&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><a href=pxzj_view.asp?action=pxzj&month="&rspxzj("month")&"&sscj="&rspxzj("sscj")&"&year="&rspxzj("year")&">"&sscjh(rspxzj("sscj"))&"</a></div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center"">"&rspxzj("year")&"年"&rspxzj("month")&"月</div></td>"
                response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><a href=tocsv.asp?action=pxzjmain&titlename=培训总结&month="&rspxzj("month")&"&sscj="&rspxzj("sscj")&"&year="&rspxzj("year")&">导出到EXCEL文档</a></div></td></tr>"
                 RowCount=RowCount-1
          rspxzj.movenext
          loop
        response.write "</table>"
       call showpage_80(page,url,total,record,PgSz)
       end if
       rspxzj.close
       set rspxzj=nothing
        conne.close
        set conne=nothing



end sub


response.write "</body></html>"



Call CloseConn
%>