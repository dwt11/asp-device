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
dim start,url,key,pagename
dim record,pgsz,total,page,rowCount,sscj,ii
action = Trim(Request("action"))
key=trim(request("keyword"))
sscj=trim(request("sscj"))

if action="keys" then 
  pagename="关键字"""&key&"&nbsp;""在 检修记录 中的搜索结果"
  url="jxjl_search.asp?keyword="&key&"&action=keys"
end if   
if action="sscjs" then 
   pagename=sscjh(sscj)&" 所有 检修记录"
   url="jxjl_search.asp?sscj="&sscj&"&action=sscjs"
end if 

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统检修记录管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>"&pagename&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""jxjl.asp"">检修记录</a>&nbsp;|&nbsp;<a href=""jxjl.asp?action=add"">添加检修记录</a>&nbsp;|&nbsp;<a href=tocsv.asp?action=jxjlmain&sql1=jxjl&titlename=检修记录>输出检修记录到Excel文档</a>"& vbCrLf
response.write " </td> </tr>"& vbCrLf
response.write "</table>"& vbCrLf
call search()
dim sqljxjl,rsjxjl
if action="keys" then sqljxjl="SELECT * from jxjl where body like '%" & key & "%' order by ID desc"
if action="sscjs" then sqljxjl="SELECT * from jxjl where sscj="&sscj&" order by ID desc"

set rsjxjl=server.createobject("adodb.recordset")
rsjxjl.open sqljxjl,conndcs,1,1
if rsjxjl.eof and rsjxjl.bof then 
response.write "<p align='center'>未添加DCS检修记录</p>" 
else

response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""8%""><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>检修原因</strong></div></td>"
response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>检修内容</strong></div></td>"
response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>检修人</strong></div></td>"
response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>检修时间</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备注</strong></div></td>"
response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
           record=rsjxjl.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsjxjl.PageSize = Cint(PgSz) 
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
           rsjxjl.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsjxjl.PageSize
           do while not rsjxjl.eof and rowcount>0
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&rsjxjl("id")&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""8%"">"&sscjh(rsjxjl("sscj"))&"</td>"
                response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rsjxjl("jxyy")&"&nbsp;</td>"
                response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&searchH(rsjxjl("body"),key)&"&nbsp;</td>"
                response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjxjl("jxrname")&"&nbsp;</div></td>"
                response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px"">"&rsjxjl("jxdate")&"&nbsp;</td>"
			    response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rsjxjl("bz")&"&nbsp;</td>"
                response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				call editdel(rsjxjl("id"),rsjxjl("sscj"),"jxjl.asp?action=editjx&id=","jxjl.asp?action=deljx&id=")
                response.write "</div></td></tr>"
                 RowCount=RowCount-1
          rsjxjl.movenext
          loop
        response.write "</table>"
       call showpage(page,url,total,record,PgSz)
       end if
       rsjxjl.close
       set rsjxjl=nothing
        conn.close
        set conn=nothing

sub search()
dim sqlcj,rscj
response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
response.write " <tr class='tdbg'><form method='Get' name='SearchForm' action='jxjl_search.asp'>" & vbCrLf
response.write "    <td>" & vbCrLf
response.write "  <strong>内容搜索：</strong>" & vbCrLf
response.write "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"" value="&key&">" & vbCrLf
response.write "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
response.write "  <input type='hidden' name='Action' value='keys'>" & vbCrLf
response.write "</td></form></strong><td><strong>查看所属车间的相关内容："
response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "	       <option value=''>按车间跳转至…</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
        response.write"<option value='jxjl_search.asp?action=sscjs&sscj="&rscj("levelid")&"'"
		if cint(request("sscj"))=rscj("levelid") then response.write" selected"
		response.write">"&rscj("levelname")&"</option>"& vbCrLf	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	response.write "     </select>	" & vbCrLf
response.write "	</td>  </tr></table>" & vbCrLf
end sub
	
Call CloseConn
%>