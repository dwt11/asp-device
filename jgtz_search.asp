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
dim sqljgtz,rsjgtz,ii,lxclassid,tyzk,zxzz
dim record,pgsz,total,page,rowCount,sscj,wcyear,sql,jx_numb
action = Trim(Request("action"))
key=trim(request("keyword"))
sscj=trim(request("sscj"))
wcyear=trim(request("wcyear"))
if action="keys" then 
  pagename="关键字"""&key&"&nbsp;""在 技改台账 中的搜索结果"
  url="jgtz_search.asp?keyword="&key&"&action=keys"
end if   
if action="sscjs" then 
   pagename=sscjh(sscj)&" 所有 技改台账"
   url="jgtz_search.asp?sscj="&sscj&"&action=sscjs"
end if 

if action="wcyears" then 
   pagename=wcyear&"完成的所有技改台账"
   url="jgtz_search.asp?wcyear="&wcyear&"&action=wcyears"
end if 


response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>信息管理系统联锁档案管理页</title>"& vbCrLf
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
response.write "    <td height='30'><a href=""jgtz.asp"">技改台账首页</a>&nbsp;|&nbsp;<a href=""jgtz.asp?action=add"">添加技改项目</a></td>"& vbCrLf
response.write "  </tr>"& vbCrLf

response.write "</table>"& vbCrLf
call search()
response.write "</table>"& vbCrLf
if action="keys" then sqljgtz="SELECT * from jgtz where name like '%" & key & "%' order by tcdate desc"
if action="sscjs" then sqljgtz="SELECT * from jgtz where sscj="&sscj&" order by tcdate desc"
if action="wcyears" then sqljgtz="SELECT * from jgtz where wc_year="&wcyear&" order by tcdate desc"
    set rsjgtz=server.createobject("adodb.recordset")
    rsjgtz.open sqljgtz,connjg,1,1
    if rsjgtz.eof and rsjgtz.bof then 
      message("未搜索到相关内容") 
    else
      record=rsjgtz.recordcount
      if Trim(Request("PgSz"))="" then
         PgSz=20
      ELSE 
         PgSz=Trim(Request("PgSz"))
      end if 
      rsjgtz.PageSize = Cint(PgSz) 
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
      rsjgtz.absolutePage = page
      start=PgSz*Page-PgSz+1
      rowCount = rsjgtz.PageSize
response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center""><strong>实施车间</strong></div></td>"
response.write "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>项目名称</strong></div></td>"
response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>提出人</strong></div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>提出时间</strong></div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>实施时间</strong></div></td>"
response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>完成时间</strong></div></td>"

response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选项</strong></div></td>"
response.write "    </tr>"
      do while not rsjgtz.eof and rowcount>0
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center"">"&rsjgtz("id")&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center"">"&sscjh(rsjgtz("sscj"))&"</div></td>"
                response.write "      <td width=""30%"" style=""border-bottom-style: solid;border-width:1px""><a href=jgtz_view.asp?id="&rsjgtz("id")&">"&searchH(rsjgtz("name"),key)&"</a>&nbsp;</div></td>"
                response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("tcr")&"&nbsp;</div></td>"
                response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("tcdate")&"&nbsp;</div></td>"
                response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("ssdate")&"&nbsp;</div></td>"
				response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsjgtz("wc_date")&"&nbsp;</div></td>"
               ' response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"
                response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"

				call editdel(rsjgtz("id"),rsjgtz("sscj"),"jgtz.asp?action=edit&id=","jgtz.asp?action=del&id=")
				
                response.write "</td></tr>"
        RowCount=RowCount-1
        rsjgtz.movenext
      loop
      response.write "</table>"
      call showpage(page,url,total,record,PgSz)
   end if
   rsjgtz.close
   set rsjgtz=nothing
   conn.close
   set conn=nothing

response.write "</body></html>"

sub search()
dim rscj,sqlcj


response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
response.write "<form method='Get' name='SearchForm' action='jgtz_search.asp'>" & vbCrLf
response.write "  <tr class='tdbg'>   <td>" & vbCrLf
response.write "  <strong>项目搜索：</strong>" & vbCrLf
response.write "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"" value="&key&">" & vbCrLf
response.write "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
response.write "  <input type='hidden' name='Action' value='keys'>" & vbCrLf
response.write "</td></form><td><font color='0066CC'> 查看所属车间的相关内容：</font>"
response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "	       <option value=''>按车间跳转至…</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='jgtz_search.asp?action=sscjs&sscj="&rscj("levelid")&"'"
		if cint(request("sscj"))=rscj("levelid") then response.write" selected"
		response.write">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	response.write "     </select>	" & vbCrLf
dwt.out "<select name='Jump2Class' id='Jump2Class' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按完成年度跳转至…</option>" & vbCrLf
	sqljgtz="SELECT distinct wc_year from jgtz"
	set rsjgtz=server.createobject("adodb.recordset")
    rsjgtz.open sqljgtz,connjg,1,1
    do while not rsjgtz.eof
		
		sql="SELECT count(id) FROM jgtz WHERE wc_year like'%"&rsjgtz("wc_year")&"%'"
		jx_numb=Connjg.Execute(sql)(0)
        
		if jx_numb<>0 then 
			Dwt.out"<option  value='jgtz_search.asp?action=wcyears&wcyear="&rsjgtz("wc_year")&"'"
			if request("wcyear")=rsjgtz("wc_year") then Dwt.out" selected"
			Dwt.out ">"&rsjgtz("wc_year")&"("&jx_numb&")</option>"& vbCrLf '
	    end if 

		rsjgtz.movenext
	loop
	rsjgtz.close
	set rsjgtz=nothing
	Dwt.out "     </select>	" & vbCrLf
	

response.write "	</td>  </tr></table>" & vbCrLf
end sub
	
Call CloseConn
%>