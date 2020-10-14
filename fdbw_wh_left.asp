<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<%
'dim sqlfdbw,rsfdbw,title,record,pgsz,total,page,start,rowcount,xh,url,ii
'dim rsadd,sqladd,fdbwid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,fdbwwh,sql,rs,czjg
dim keys,sscjid,ssghid

fdbwid=Trim(Request("fdbwid"))
'fdbwwh=trim(request("fdbwwh"))	
keys=trim(request("keyword")) 
sscjid=trim(request("sscj"))
ssghid=trim(request("ssgh")) 
dwt.out"<html>"& vbCrLf
dwt.out"<head>" & vbCrLf
dwt.out"<title>防冻保温检修记录汇总</title>"& vbCrLf
dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"</head>"& vbCrLf
dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>防冻保温检修记录汇总</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

    call search()

if request("action")="" then call main
'if request("action")="jxcount" then call jxcount


sub main()
dim fdbwid,sqlfdbw,rsfdbw
dim record,pgsz,total,page,start,rowcount,url
dim title

sqlfdbw = "SELECT fdbw_whjl.*, fdbw.sscj,fdbw.wh,fdbw.ssgh FROM fdbw_whjl INNER JOIN fdbw ON fdbw_whjl.fdbwid = fdbw.id "

if keys<>"" then 
		sqlfdbw=sqlfdbw&"WHERE (((fdbw.wh) like '%" &keys& "%')) "
		title="-搜索 "&keys
		url="fdbw_wh_left.asp?keyword="&keys
	end if 
if sscjid<>"" then
		sqlfdbw=sqlfdbw&"WHERE (((fdbw.sscj)="&sscjid&")) "
		title="-"&sscjh(sscjid)
		url="fdbw_wh_left.asp?sscj="&sscjid
	end if 
if ssghid<>"" then
        sqlfdbw=sqlfdbw&"WHERE (((fdbw.ssgh)="&ssghid&")) "
	    title="-"&gh(ssghid)
		url="fdbw_wh_left.asp?ssgh="&ssghid
	end if 

  sqlfdbw=sqlfdbw&"order by  whsj DESC"	

set rsfdbw=server.createobject("adodb.recordset")
rsfdbw.open sqlfdbw,connjg,1,1

if rsfdbw.eof and rsfdbw.bof then 
message("没有检修记录")

else
       record=rsfdbw.recordcount
		
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		
		rsfdbw.PageSize = Cint(PgSz) 
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
		rsfdbw.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rsfdbw.PageSize
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out "<tr class=""x-grid-header"">" & vbCrLf
dwt.out  "     <td  width=""5%"" class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
dwt.out  "      <td  width=""8%"" class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
dwt.out  "      <td  width=""5%"" class='x-td'><DIV class='x-grid-hd-text'>工号</div></td>"& vbCrLf
dwt.out  "      <td  width=""15%"" class='x-td'><DIV class='x-grid-hd-text'>位号</div></td>"& vbCrLf
dwt.out  "      <td  width=""12%"" class='x-td'><DIV class='x-grid-hd-text'>维护原因</div></td>"& vbCrLf
dwt.out  "      <td  width=""10%"" class='x-td'><DIV class='x-grid-hd-text'>维护时间</div></td>"& vbCrLf
dwt.out  "      <td  width=""35%"" class='x-td'><DIV class='x-grid-hd-text'>维护内容</div></td>"& vbCrLf
dwt.out  "      <td  width=""10%"" class='x-td'><DIV class='x-grid-hd-text'>维护结果</div></td>"& vbCrLf

dwt.out  "    </tr>"

           do while not rsfdbw.eof and rowcount>0
		   dim tyqk
				select case rsfdbw("whjg")
			  case 1
				 tyqk="<span style='color:#006600'>投运</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>具备条件</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>有缺陷</span>"
			  case 4 
				tyqk="保温取消"
			end select	 
		   xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if  
			
			
		dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_d(rsfdbw("sscj"))&"</div></td>"& vbCrLf
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&gh(rsfdbw("ssgh"))&"&nbsp;</td>" & vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><a href=fdbw_whjl.asp?fdbwid="&rsfdbw("fdbwid")&">"&searchH(uCase(rsfdbw("wh")),keys)&"</div></td>"& vbCrLf
			
			dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rsfdbw("whyy")&"&nbsp;</td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("whsj")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfdbw("body")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&tyqk&"&nbsp;</div></td>"
              
     
                dwt.out  "</div></td></tr>"	
			RowCount=RowCount-1
          rsfdbw.movenext
          loop
        dwt.out  "</table>"
			
		if sscjid<>"" or keys<>""or ssghid<>"" then 
				call showpage(page,url,total,record,PgSz)
        else
				call showpage1(page,url,total,record,PgSz)
		end if 
			dwt.out "</div>"& vbCrLf
    end if
	dwt.out "</div>"  
	rsfdbw.close
       set rsfdbw=nothing
        connjg.close
        set connjg=nothing
	
end sub	

sub search()
	dim sqlcj,rscj,sqlgh,rsgh
	dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='fdbw_wh_left.asp'>" & vbCrLf
	'if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then  dwt.out "<a href='fdbw.asp?action=add'>添加防冻保温</a>"
	'dwt.out "&nbsp;&nbsp;<a href='fdbw.asp?update=update'>查看最近七天更新</a>"
	dwt.out "  <input type='text' name='keyword'  size='20' maxlength='50' "
	if keys<>"" then 
	 dwt.out "value='"&keys&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='输入搜索的位号'"
	 	dwt.out " onblur=""if(this.value==''){this.value='输入搜索的位号'}"" onfocus=""this.value=''"">" & vbCrLf
	end if                 
	dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	dwt.out "&nbsp;&nbsp;<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='fdbw_wh_left.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid")  then dwt.out" selected"
			dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf

	dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>按装置跳转至…</option>" & vbCrLf
	sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
		set rsgh=server.createobject("adodb.recordset")
		rsgh.open sqlgh,conn,1,1
		do while not rsgh.eof
			dwt.out"<option value='fdbw_wh_left.asp?ssgh="&rsgh("ghid")&"'"
			if cint(request("ssgh"))=rsgh("ghid") then dwt.out" selected"
			dwt.out ">"&rsgh("gh_name")&"("&Connjg.Execute("SELECT count(id) FROM fdbw WHERE ssgh="&rsgh("ghid"))(0)&")</option>"& vbCrLf
		
			rsgh.movenext
		loop
		rsgh.close
		set rsgh=nothing
		dwt.out "     </select>	" & vbCrLf
		dwt.out "</form></div></div>" & vbCrLf

end sub

dwt.out "</body></html>"
Call Closeconn













%>
