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
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统安全活动月计划总结页</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  " <tr class='topbg'>"& vbCrLf
dwt.out  "   <td height='22' colspan='2' align='center'><strong>安全活动月计划总结页</strong></td>"& vbCrLf
dwt.out  "  </tr>  "& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
dwt.out  "    <td height='30'><a href=""aqyjhzj.asp"">安全活动月计划总结首页</a>&nbsp;|&nbsp;<a href=""aqyjh_view.asp?action=addyjh"">添加月计划</a>&nbsp;|&nbsp;<a href=""aqyzj_view.asp?action=addyzj"">添加月总结</a></td>"& vbCrLf
dwt.out  "  </tr>"& vbCrLf
dwt.out  "</table>"& vbCrLf

if request("action")="yjh_bz" then 
	call yjh_bz
else  
	if request("action")="yzj_bz" then
	   call yzj_bz
    else
	   call main 
	end if   
end if 	  

'*****************************************************
'列表显示月份
'**********************************************88888888888
sub main()
  dim i,ii
  dim sql,rs,years(100),months(100)
  ii=1
   
   
   '显示月计划
   dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""<tr  class=""title""><td height=30 style=""border-bottom-style: solid;border-width:1px"" colspan=""3""><div align=center>安全活动月计划</div></td></tr><tr class='tdbg'><td>"
   sql="SELECT distinct year,month from yjh "
   set rs=server.createobject("adodb.recordset")
   rs.open sql,connaq,1,1
      if rs.eof and rs.bof then
      dwt.out  "<div align=center><font color=#00000>没有添加月计划</font></div>"
  else

   
   do while not rs.eof
     i=i+1
     
     years(i)=rs("year")
	 months(i)=rs("month")
     
	  RS.movenext
      loop
   end if
   rs.close
   set rs=nothing
  
   for i=i to 1 step -1
	 if ii>8 then 
	  dwt.out  "<br>"
	  ii=1
	 end if
	 ii=ii+1
	 if len(months(i))<>2 then months(i)="0"&months(i)  
	 dwt.out  "&nbsp;&nbsp;&nbsp;&nbsp;<a href=aqyjhzj.asp?action=yjh_bz&year="&years(i)&"&month="&months(i)&">"&years(i)&"年"&months(i)&"月</a>&nbsp;&nbsp;&nbsp;"
   next
   dwt.out "</tr></td></table>"
   
   
   
      '显示月总结
   dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""<tr  class=""title""><td height=30 style=""border-bottom-style: solid;border-width:1px"" colspan=""3""><div align=center>安全活动月总结</div></td></tr><tr class='tdbg'><td>"
   sql="SELECT distinct year,month from yzj "
   set rs=server.createobject("adodb.recordset")
   rs.open sql,connaq,1,1
   if rs.eof and rs.bof then
      dwt.out  "<div align=center><font color=#00000>没有添加月总结</font></div>"
  else
   do while not rs.eof
     i=i+1
     
     years(i)=rs("year")
	 months(i)=rs("month")
     
	  RS.movenext
      loop
  end if 	  
   rs.close
   set rs=nothing
  
   for i=i to 1 step -1
	 if ii>8 then 
	  dwt.out  "<br>"
	  ii=1
	 end if
	 ii=ii+1
	 if len(months(i))<>2 then months(i)="0"&months(i)  
	 dwt.out  "&nbsp;&nbsp;&nbsp;&nbsp;<a href=aqyjhzj.asp?action=yzj_bz&year="&years(i)&"&month="&months(i)&">"&years(i)&"年"&months(i)&"月</a>&nbsp;&nbsp;&nbsp;"
   next
   dwt.out "</tr></td></table>"

   
end sub

'*****************************************************
'列表每个月各车间报的月计划,点击月份后显示
'**********************************************88888888888
sub yjh_bz()    
dim xh
   dwt.out  "<div align=center>"&request("year")&"年"&request("month")&"安全活动月工作计划</div>"
   dim sqlyjh,rsyjh
   sqlyjh="SELECT * from yjh where month="&request("month")&" and year="&request("year")
   set rsyjh=server.createobject("adodb.recordset")
   rsyjh.open sqlyjh,connaq,1,1
   if rsyjh.eof and rsyjh.bof then 
      dwt.out  "<p align='center'>未添加月计划</p>" 
   else
      dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      dwt.out  "<tr class=""title"">" 
      dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
      dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><strong>单&nbsp;&nbsp;&nbsp;&nbsp;位</strong></div></td>"
      dwt.out  "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选&nbsp;项</strong></div></td>"
      dwt.out  "    </tr>"
      do while not rsyjh.eof
		xh=xh+1
                dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
				dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><a href=aqyjh_view.asp?action=yjh&month="&request("month")&"&sscj="&rsyjh("sscj")&"&year="&request("year")&">"&sscjh(rsyjh("sscj"))&"</a></div></td>"
				dwt.out  "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;"
				'<a href=tocsv.asp?action=yjhmain&titlename=月计划&month="&request("month")&"&sscj="&rsyjh("sscj")&"&year="&request("year")&">导出到EXCEL文档</a>
				call editdel(rsyjh("id"),rsyjh("sscj"),"aqyjh_view.asp?action=edit&id=","aqyjh_view.asp?action=del&id=")
                dwt.out  "</div></td></tr>"
          rsyjh.movenext
     loop
     dwt.out  "</table>"
 end if
       rsyjh.close
       set rsyjh=nothing
end sub

'*****************************************************
'列表每个月各车间报的月总结,点击月份后显示
'**********************************************88888888888
sub yzj_bz()    
dim xh
   dwt.out  "<div align=center>"&request("year")&"年"&request("month")&"安全活动月工作总结</div>"
   dim sqlyjh,rsyzj
   sqlyjh="SELECT * from yzj where month="&request("month")&" and year="&request("year")
   set rsyzj=server.createobject("adodb.recordset")
   rsyzj.open sqlyjh,connaq,1,1
   if rsyzj.eof and rsyzj.bof then 
      dwt.out  "<p align='center'>未添加月总结</p>" 
   else
      dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      dwt.out  "<tr class=""title"">" 
      dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>序号</strong></div></td>"
      dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><strong>单&nbsp;&nbsp;&nbsp;&nbsp;位</strong></div></td>"
      dwt.out  "      <td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选&nbsp;项</strong></div></td>"
      dwt.out  "    </tr>"
      do while not rsyzj.eof
		xh=xh+1
                dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
				dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><a href=aqyzj_view.asp?action=yzj&month="&request("month")&"&sscj="&rsyzj("sscj")&"&year="&request("year")&">"&sscjh(rsyzj("sscj"))&"</a></div></td>"
				dwt.out  "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				'<a href=tocsv.asp?action=yzjmain&titlename=月总结&month="&request("month")&"&sscj="&rsyzj("sscj")&"&year="&request("year")&">导出到EXCEL文档</a>
                call editdel(rsyzj("id"),rsyzj("sscj"),"aqyzj_view.asp?action=edit&id=","aqyzj_view.asp?action=del&id=")
				dwt.out  "</div></td></tr>"
          rsyzj.movenext
     loop
     dwt.out  "</table>"
 end if
       rsyzj.close
       set rsyzj=nothing
end sub



dwt.out  "</body></html>"



Call CloseConn
%>