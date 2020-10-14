<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<%
'dim sqllsda,rslsda,title,record,pgsz,total,page,start,rowcount,xh,url,ii
'dim rsadd,sqladd,lsdaid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,lsdawh,sql,rs,czjg
dim keys,sscjid,ssghid

lsdaid=Trim(Request("lsdaid"))
'lsdawh=trim(request("lsdawh"))	
keys=trim(request("keyword")) 
sscjid=trim(request("sscj"))
ssghid=trim(request("ssgh")) 
dwt.out"<html>"& vbCrLf
dwt.out"<head>" & vbCrLf
dwt.out"<title>联锁检修记录汇总</title>"& vbCrLf
dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"</head>"& vbCrLf
dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>联锁检修记录汇总</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

    call search()

if request("action")="" then call main
'if request("action")="jxcount" then call jxcount


sub main()
dim lsdaid,sqllsda,rslsda
dim record,pgsz,total,page,start,rowcount,url
dim title

sqllsda = "SELECT lsda_czjl.*, lsda.sscj,lsda.wh,lsda.ssgh FROM lsda_czjl INNER JOIN lsda ON lsda_czjl.lsdaid = lsda.lsdaid "

if keys<>"" then 
		sqllsda=sqllsda&"WHERE (((lsda.wh) like '%" &keys& "%')) "
		title="-搜索 "&keys
		url="lsda_wh_left.asp?keyword="&keys
	end if 
if sscjid<>"" then
		sqllsda=sqllsda&"WHERE (((lsda.sscj)="&sscjid&")) "
		title="-"&sscjh(sscjid)
		url="lsda_wh_left.asp?sscj="&sscjid
	end if 
if ssghid<>"" then
        sqllsda=sqllsda&"WHERE (((lsda.ssgh)="&ssghid&")) "
	    title="-"&gh(ssghid)
		url="lsda_wh_left.asp?ssgh="&ssghid
	end if 

  sqllsda=sqllsda&"order by  czsj DESC"	

set rslsda=server.createobject("adodb.recordset")
rslsda.open sqllsda,connjg,1,1

if rslsda.eof and rslsda.bof then 
message("没有检修记录")

else
       record=rslsda.recordcount
		
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		
		rslsda.PageSize = Cint(PgSz) 
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
		rslsda.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rslsda.PageSize
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out "<tr class=""x-grid-header"">" & vbCrLf
dwt.out  "     <td  width=""5%"" class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
dwt.out  "      <td  width=""8%"" class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
dwt.out  "      <td  width=""5%"" class='x-td'><DIV class='x-grid-hd-text'>工号</div></td>"& vbCrLf
dwt.out  "      <td  width=""15%"" class='x-td'><DIV class='x-grid-hd-text'>位号</div></td>"& vbCrLf
dwt.out  "      <td  width=""12%"" class='x-td'><DIV class='x-grid-hd-text'>操作原因</div></td>"& vbCrLf
dwt.out  "      <td  width=""10%"" class='x-td'><DIV class='x-grid-hd-text'>操作时间</div></td>"& vbCrLf
dwt.out  "      <td  width=""10%"" class='x-td'><DIV class='x-grid-hd-text'>操作结果</div></td>"& vbCrLf

dwt.out  "    </tr>"

           do while not rslsda.eof and rowcount>0
		   dim czjg
        select case rslsda("czjg")
          case 0
            czjg="旁路"
           if rslsda("czyy") then
		    czjg="<font color='#ff0000'>"&czjg&"</font>"
		   else
		    czjg="<font color='#0000ff'>"&czjg&"</font>"
		   end if 	
		  case 1 
        	czjg="投运"
        end select	 
		   xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if  
			
			
		dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_d(rslsda("sscj"))&"</div></td>"& vbCrLf
			 dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&gh(rslsda("ssgh"))&"&nbsp;</td>" & vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><a href=lsda_whjl.asp?lsdaid="&rslsda("lsdaid")&">"&searchH(uCase(rslsda("wh")),keys)&"</div></td>"& vbCrLf
			
			dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rslsda("czinfo")&"&nbsp;</td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rslsda("czsj")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&czjg&"&nbsp;</div></td>"
              
     
                dwt.out  "</div></td></tr>"	
			RowCount=RowCount-1
          rslsda.movenext
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
	rslsda.close
       set rslsda=nothing
        connjg.close
        set connjg=nothing
	
end sub	

sub search()
	dim sqlcj,rscj,sqlgh,rsgh
	dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='ls_wh_left.asp'>" & vbCrLf
	'if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then  dwt.out "<a href='lsda.asp?action=add'>添加防冻保温</a>"
	'dwt.out "&nbsp;&nbsp;<a href='lsda.asp?update=update'>查看最近七天更新</a>"
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
			dwt.out"<option value='ls_wh_left.asp?sscj="&rscj("levelid")&"'"
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
			dwt.out"<option value='ls_wh_left.asp?ssgh="&rsgh("ghid")&"'"
			if cint(request("ssgh"))=rsgh("ghid") then dwt.out" selected"
			dwt.out ">"&rsgh("gh_name")&")</option>"& vbCrLf
		
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
