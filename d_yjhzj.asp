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
dwt.out  "<title>信息管理系统月计划总结页</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  " <tr class='topbg'>"& vbCrLf
dwt.out  "   <td height='22' colspan='2' align='center'><strong>党委月计划总结页</strong></td>"& vbCrLf
dwt.out  "  </tr>  "& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
dwt.out  "    <td height='30'><a href=""d_yjhzj.asp"">月计划总结首页</a>&nbsp;|&nbsp;<a href=""d_yjh_view.asp?action=addyjh"">添加月计划</a>&nbsp;|&nbsp;<a href=""d_yzj_view.asp?action=addyzj"">添加月总结</a></td>"& vbCrLf
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
   dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""<tr  class=""title""><td height=30 style=""border-bottom-style: solid;border-width:1px"" colspan=""3""><div align=center>月计划</div></td></tr><tr class='tdbg'><td>"
   sql="SELECT distinct year,month from d_yjh "
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conna,1,1
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
	 dwt.out  "&nbsp;&nbsp;&nbsp;&nbsp;<a href=d_yjhzj.asp?action=yjh_bz&year="&years(i)&"&month="&months(i)&">"&years(i)&"年"&months(i)&"月</a>&nbsp;&nbsp;&nbsp;"
   next
   dwt.out "</tr></td></table>"
   
   
   dim sql1,rs1
      '显示月总结
   dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""<tr  class=""title""><td height=30 style=""border-bottom-style: solid;border-width:1px"" colspan=""3""><div align=center>月总结</div></td></tr><tr class='tdbg'><td>"
   sql1="SELECT distinct year,month from d_yzj "
   set rs1=server.createobject("adodb.recordset")
   rs1.open sql1,conna,1,1
   if rs1.eof and rs1.bof then
      dwt.out  "<div align=center><font color=#00000>没有添加月总结</font></div>"
  else
   do while not rs1.eof
     i=i+1
     
     years(i)=rs1("year")
	 months(i)=rs1("month")
     
	  RS1.movenext
      loop
  end if 	  
   rs1.close
   set rs=nothing
  
   for i=i to 1 step -1
	 if ii>8 then 
	  dwt.out  "<br>"
	  ii=1
	 end if
	 ii=ii+1
	 if len(months(i))<>2 then months(i)="0"&months(i)  
	 dwt.out  "&nbsp;&nbsp;&nbsp;&nbsp;<a href=d_yjhzj.asp?action=yzj_bz&year="&years(i)&"&month="&months(i)&">"&years(i)&"年"&months(i)&"月</a>&nbsp;&nbsp;&nbsp;"
   next
   dwt.out "</tr></td></table>"

   
end sub

'*****************************************************
'列表每个月各车间报的月计划,点击月份后显示
'**********************************************88888888888
sub yjh_bz()    
dim xh
   dwt.out  "<div align=center>"&request("year")&"年"&request("month")&"月份工作计划</div>"
   dim sqlyjh,rsyjh
   sqlyjh="SELECT * from d_yjh where month="&request("month")&" and year="&request("year")
   set rsyjh=server.createobject("adodb.recordset")
   rsyjh.open sqlyjh,conna,1,1
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
		dim sszb
		if rsyjh("sscj")=1 then sszb="维修一党织部"
		if rsyjh("sscj")=2 then sszb="维修二党织部"
		if rsyjh("sscj")=3 then sszb="机关党织部"
       		if rsyjh("sscj")=4 then sszb="维修三党支部"
		if rsyjh("sscj")=5 then sszb="维修四党支部"

	            dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
				dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><a href=d_yjh_view.asp?action=yjh&month="&request("month")&"&sscj="&rsyjh("sscj")&"&year="&request("year")&">"&sszb&"</a></div></td>"
				dwt.out  "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;"
				'<a href=tocsv.asp?action=yjhmain&titlename=月计划&month="&request("month")&"&sscj="&rsyjh("sscj")&"&year="&request("year")&">导出到EXCEL文档</a>
                if rsyjh("userid")=session("userid") then  response.Write "<a href=d_yjh_view.asp?action=edit&id="&rsyjh("id")&">编辑</a> <a href=d_yjh_view.asp?action=del&id="&rsyjh("id")&">删除</a>"
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
   dwt.out  "<div align=center>"&request("year")&"年"&request("month")&"月份工作总结</div>"
   dim sqlyjh,rsyzj
   sqlyjh="SELECT * from d_yzj where month="&request("month")&" and year="&request("year")
   set rsyzj=server.createobject("adodb.recordset")
   rsyzj.open sqlyjh,conna,1,1
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
 		dim sszb
		if rsyzj("sscj")=1 then sszb="维修一党织部"
		if rsyzj("sscj")=2 then sszb="维修二党织部"
		if rsyzj("sscj")=3 then sszb="机关党织部"
               		if rsyzj("sscj")=4 then sszb="维修三党支部"
		if rsyzj("sscj")=5 then sszb="维修四党支部"

       dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&xh&"</div></td>"
				dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""63%""><div align=""center""><a href=d_yzj_view.asp?action=yzj&month="&request("month")&"&sscj="&rsyzj("sscj")&"&year="&request("year")&">"&sszb&"</a></div></td>"
				dwt.out  "<td width=""20%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				'<a href=tocsv.asp?action=yzjmain&titlename=月总结&month="&request("month")&"&sscj="&rsyzj("sscj")&"&year="&request("year")&">导出到EXCEL文档</a>
                if rsyzj("userid")=session("userid") then response.Write  "<a href=d_yzj_view.asp?action=edit&id="&rsyzj("id")&">编辑</a> <a href=d_yzj_view.asp?action=del&id="&rsyzj("id")&">删除</a>"
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