<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->


<%
dim url,record,pgsz,total,page,start,rowcount,ii,pagename
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统操作日志管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf


action=request("action")
	dim leftmdb,connleft,connl
	leftmdb="ybdata/left.mdb"
	Set connleft = Server.CreateObject("ADODB.Connection")
	connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
	connleft.Open connl    

select case action
'  case "add"
'    if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
'  case "saveadd"
'    call saveadd
'  case "edit"
'	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
'  case "saveedit"
'    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	


sub del()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from systemlog where id="&id
  rsdel.open sqldel,connleft,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub


sub main()
    url="systemlog.asp"
	dim sql,rs,title
	sql="SELECT * from systemlog"
'	if keys<>"" then 
'		sql=sql&" where body like '%" &keys& "%' "
'		title="-搜索 "&keys
'	end if 
'	if sscjid<>"" then
'		sql=sql&" where sscj="&sscjid
'		title="-"&sscjh(sscjid)
'	end if 
	sql=sql&" ORDER BY update desc"
	
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>系统操作记录</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	
'	'for sscji=1 to 5 '071017修改
'	sql="select * from levelname where istq=false"
'	set rs=server.createobject("adodb.recordset")
'	rs.open sql,connleft,1,1
'	if rs.eof and rs.bof then 
'		dwt.out "没有添加车间"
'	else
'	   do while not rs.eof
'		sql="SELECT count(id) FROM jxjl WHERE sscj="&rs("levelid")&" and month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
'		numb=numb&sscjh_d(rs("levelid"))&":"&"<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
'	rs.movenext
'	loop
'	end if 
'	rs.close
'	
'	sql="SELECT count(id) FROM jxjl WHERE  month(jxdate)="&month(now)&"and year(jxdate)="&year(now())
'	totall= "<span style='color:#006600;'>"&conndcs.Execute(sql)(0)&"</span>" 
'	dwt.out "<div class='pre'>本月"&numb&"合计:"&totall&"</div>"& vbCrLf
'
'	search()

	set rs=server.createobject("adodb.recordset")
	rs.open sql,connleft,1,1
	if rs.eof and rs.bof then 
		message("未找到相关检修记录")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>内容</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>用户名</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>时间</div></td>"& vbCrLf
		dwt.out "      <td   class='x-td'><DIV class='x-grid-hd-text'>IP</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"& vbCrLf
		dwt.out "    </tr>"& vbCrLf
		record=rs.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rs.PageSize = Cint(PgSz) 
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
		rs.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rs.PageSize
		do while not rs.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rs("action")&" "&rs("leftname")&" "&rs("message")&"</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&usernameh(rs("userid"))&"("&useridh(rs("userid"))&")</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rs("update")&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rs("ip")&"</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center><a href='systemlog.asp?action=del&id="&rs("id")&"'  onClick=""return confirm('确定要删除此记录吗？');"">删除</a></div></td></tr>"& vbCrLf
			RowCount=RowCount-1
			rs.movenext
		loop
		dwt.out "</table>"& vbCrLf
		  call showpage1(page,url,total,record,PgSz)
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end sub

dwt.out "</body></html>"
%>