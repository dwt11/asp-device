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
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql

'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
dim pagename

if Request("action")="xc" then pagename="现存台账---关键字“"&request("keyword")&"”"
if Request("action")="sr" then 
    if request("keyword")="" then pagename="入库台账---关键字“”"
	if request("keyword")<>"" then pagename="入库台账---关键字“"&request("keyword")&"”"
    if request("qsdate")<>"" then pagename="入库台账---“"&request("qsdate")&"”到“"&request("zzdate")&"”记录"
end if 
if Request("action")="fc" then 
        if request("keyword")="" then pagename="出库台账---关键字“”"
    if request("keyword")<>"" then pagename="出库台账---关键字“"&request("keyword")&"”"
    if request("qsdate")<>"" then pagename="出库台账---“"&request("qsdate")&"”到“"&request("zzdate")&"”记录"
end if 

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>库存台账管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>库存台账搜索页---"&pagename&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='10%'><strong>管理导航：</strong></td>"& vbCrLf
response.write "    <td height='90%'><strong><a href=kcgl.asp>现存</a>&nbsp;&nbsp;<a href=kcgl_sr.asp>入库</a>&nbsp;&nbsp;<a href=kcgl_fc.asp>出库</a></strong>"
response.write "</td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

if Request("action")="xc" then call xc '主页面显示最新库存信息
if Request("action")="sr" then call sr '主页面显示最新入库存信息
if Request("action")="fc" then call fc '主页面显示最新出库存信息

Sub xc()
dim sqlbody,rsbody,xh
if request("keyword")="" then 
   url="kcgl_search.asp?action=xc"
   sqlbody="SELECT * from xc order by id DESC"
end if 
if request("keyword")<>"" then 
   url="kcgl_search.asp?action=xc&keyword="&request("keyword")
   sqlbody="SELECT * from xc where name like '%" & request("keyword") & "%' order by id DESC"
end if 
  'sqlbody="SELECT * from xc order by id DESC"
  on error resume next
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">暂无内容</p>" 
  else
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
  
     response.write "<table   border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" width=""100%"">"
     response.write "<tr class=""title"">"
     response.write "<td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>编号</strong></div></td>"
     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>车间</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>分类</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>名称</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>规格型号</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单位</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单价</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>数量</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>金额</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>入库时间</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备 注</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选 项</strong></div></td>"
     response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        xh=xh+1
        'on error resume next
		response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("wpid")&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sscj"))&"</div></td>"

		response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&dclass(rsbody("class"))&"-"&kcclass(rsbody("class"))&"</div></td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&searchH(rsbody("name"),request("keyword"))&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&rsbody("xhgg")&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dw")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dmoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("numb")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("amoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("rcdate")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("bz")&"&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
       call editdel(rsbody("id"),rsbody("sscj"))
	   response.write "</div></td></tr>"
       dim totalamoney '合计页里的总金额
	   totalamoney=totalamoney+rsbody("amoney")
	    RowCount=RowCount-1
    rsbody.movenext
    loop
   
   
           response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color=#FF0000>合计</font></div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color=#FF0000>"&totalamoney&"</font>&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td></tr>"

   response.write "</table>"
     call showpage(page,url,total,record,PgSz)
 end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub









sub sr()
dim sqlbody,rsbody,xh


if request("keyword")="" and request("qsdate")="" then 
   url="kcgl_search.asp?action=sr"
   sqlbody="SELECT * from sr order by id DESC"
end if 
if request("keyword")<>"" then 
   url="kcgl_search.asp?action=sr&keyword="&request("keyword")
   sqlbody="SELECT * from sr where name like '%" & request("keyword") & "%' order by id DESC"
end if 

if request("qsdate")<>"" then 
   url="kcgl_search.asp?action=sr&qsdate="&request("qsdate")&"&zzdate="&request("zzdate")
   sqlbody="SELECT * from sr where srdate between #"&request("qsdate")&"# and #"&request("zzdate")&"# order by id DESC"
end if 

  on error resume next
  'sqlbody="SELECT * from sr order by id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">暂无内容</p>" 
  else
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
  
     response.write "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  width=""100%"">"
     response.write "<tr class=""title"">"
     response.write "<td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>编号</strong></div></td>"
     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>车间</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>分类</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>来源</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>名称</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>规格型号</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单位</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单价</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>数量</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>金额</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>入库时间</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备 注</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选 项</strong></div></td>"
     response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        xh=xh+1
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("wpid")&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sscj"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&dclass(rsbody("class"))&"-"&kcclass(rsbody("class"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("lytxt")&"&nbsp;</div></td>"

        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&searchH(rsbody("name"),request("keyword"))&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&rsbody("xhgg")&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dw")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dmoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("numb")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("amoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("sr_year")&"-"&rsbody("sr_month")&"-"&rsbody("sr_day")&"</div></td>"
		response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("bz")&"&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
       if session("level")=rsbody("sscj") or session("level")=0 then 
	    response.write "<a href=kcgl_sr.asp?action=delsrinfo&id="&rsbody("id")&" onClick=""return confirm('确定要删除此入库记录吗？');"">删除</a>"
     else
        response.write "&nbsp;"
     end if 
	   response.write "</div></td></tr>"
       dim totalamoney '合计页里的总金额
	   totalamoney=totalamoney+rsbody("amoney")
	    RowCount=RowCount-1
    rsbody.movenext
    loop
           response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color=#FF0000>合计</font></div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
       response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color=#FF0000>"&totalamoney&"</font>&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"

       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td></tr>"

   response.write "</table>"
     call showpage(page,url,total,record,PgSz)
  end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
END SUB

sub fc()
dim sqlbody,rsbody,xh
if request("keyword")="" and request("qsdate")="" then 
   url="kcgl_search.asp?action=fc"
   sqlbody="SELECT * from fc order by id DESC"
end if 
if request("keyword")<>"" then 
   url="kcgl_search.asp?action=fc&keyword="&request("keyword")
   sqlbody="SELECT * from fc where name like '%" & request("keyword") & "%' order by id DESC"
end if 

if request("qsdate")<>"" then 
   url="kcgl_search.asp?action=sr&qsdate="&request("qsdate")&"&zzdate="&request("zzdate")
   sqlbody="SELECT * from fc where srdate between #"&request("qsdate")&"# and #"&request("zzdate")&"# order by id DESC"
end if 

  'on error resume next
  'sqlbody="SELECT * from fc order by id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,connkc,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">暂无内容</p>" 
  else
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
  
     response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     response.write "<tr class=""title"">"
     response.write "<td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>编号</strong></div></td>"
     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>车间</strong></div></td>"
     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>分类</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>去向</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>名称</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>规格型号</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单位</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单价</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>数量</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>金额</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>出库时间</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备 注</strong></div></td>"
     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选 项</strong></div></td>"
     response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
  xh=xh+1
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("wpid")&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sscj"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&dclass(rsbody("class"))&"-"&kcclass(rsbody("class"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"
		if rsbody("qxtxt")=1000 then 
		   response.write "现场使用"
		else
		   response.write sscjh(rsbody("qxtxt"))
        end if 
		response.write "</div></td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&searchH(rsbody("name"),request("keyword"))&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">"&rsbody("xhgg")&"&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dw")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("dmoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("numb")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("amoney")&"&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("fc_year")&"-"&rsbody("fc_month")&"-"&rsbody("fc_day")&"</div></td>"
		response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("bz")&"&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
       if session("level")=rsbody("sscj") or session("level")=0 then 
	    response.write "<a href=kcgl_fc.asp?action=delfcinfo&id="&rsbody("id")&" onClick=""return confirm('确定要删除此出库记录吗？');"">删除</a>"
     else
        response.write "&nbsp;"
     end if 
	   response.write "</div></td></tr>"
       dim totalamoney '合计页里的总金额
	   totalamoney=totalamoney+rsbody("amoney")
	    RowCount=RowCount-1
    rsbody.movenext
    loop
           response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><font color=#FF0000>合计</font></div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" >&nbsp;</td>"
       response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px"">&nbsp;</td>"
        response.write "  <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><font color=#FF0000>"&totalamoney&"</font>&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>"
       response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td></tr>"
   response.write "</table>"
     call showpage(page,url,total,record,PgSz)
 	  end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub 

response.write "</body></html>"








'用于库存子分类名称显示
Function kcclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from kcclass where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connkc,1,1
    do while not rsname.eof
	    kcclass=rsname("name")
		rsname.movenext
	loop
	rsname.close
	set rsname=nothing
end Function 

'用于显示父分类名称 
Function dclass(classid)
	dim sqlname,rsname
	dim sqlz,rsz
	sqlz="SELECT * from kcclass where id="&classid
    set rsz=server.createobject("adodb.recordset")
    rsz.open sqlz,connkc,1,1
    'do while not rsz.eof
	 '   kcclass=rsname("name")
		'rsname.movenext
	'loop
	   sqlname="SELECT * from class where id="&rsz("class")
       set rsname=server.createobject("adodb.recordset")
       rsname.open sqlname,connkc,1,1
       'do while not rsname.eof
	    dclass=rsname("name")
		'rsname.movenext
	'loop
	rsname.close
	set rsname=nothing
	rsz.close
	set rsz=nothing
end Function 



'选项（编辑、出库\删除）
sub editdel(id,sscj)
 if session("level")=sscj or session("level")=0 then 
	response.write "<a href=kcgl_fcsa.asp?id="&id&">出库</a>&nbsp;"
	response.write "<a href=kcgl.asp?action=del&id="&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 else
    response.write "&nbsp;"
 end if 
end sub


Call CloseConn
%>