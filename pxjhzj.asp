<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
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
response.write "<title>��Ϣ����ϵͳ��ѵ����ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
if request("action")="pxjh" then call pxjh
if request("action")="pxzj" then call pxzj


sub pxjh()
response.write "<br><br><table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>��ѵ�ƻ�ҳ</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf


url="pxjhzj.asp?action=pxjh"
sqlpxjh="SELECT  distinct sscj,year,month  from pxjh order by year desc,month  DESC"
set rspxjh=server.createobject("adodb.recordset")
rspxjh.open sqlpxjh,conne,1,1
if rspxjh.eof and rspxjh.bof then 
message("δ�����ѵ�ƻ�")
else

response.write "<table  width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center""><strong>���</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>ѡ��</strong></div></td>"
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
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center"">"&rspxjh("year")&"��"&rspxjh("month")&"��</div></td>"
                response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><a href=tocsv.asp?action=pxjhmain&titlename=��ѵ�ƻ�&month="&rspxjh("month")&"&sscj="&rspxjh("sscj")&"&year="&rspxjh("year")&">������EXCEL�ĵ�</a></div></td></tr>"
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
response.write "   <td height='22' colspan='2' align='center'><strong>��ѵ�ܽ�ҳ</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf
url="pxjhzj.asp?action=pxzj"
sqlpxzj="SELECT  distinct sscj,year,month  from pxzj order by year desc, month  DESC"
set rspxzj=server.createobject("adodb.recordset")
rspxzj.open sqlpxzj,conne,1,1
if rspxzj.eof and rspxzj.bof then 
message("δ�����ѵ�ܽ�")
else

response.write "<table  width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""10%""><div align=""center""><strong>���</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""50%""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>����</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><strong>ѡ��</strong></div></td>"
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
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center"">"&rspxzj("year")&"��"&rspxzj("month")&"��</div></td>"
                response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""20%""><div align=""center""><a href=tocsv.asp?action=pxzjmain&titlename=��ѵ�ܽ�&month="&rspxzj("month")&"&sscj="&rspxzj("sscj")&"&year="&rspxzj("year")&">������EXCEL�ĵ�</a></div></td></tr>"
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