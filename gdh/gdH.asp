<%@language=vbscript codepage=936 %>
<%

'gdh.asp为YB007上用户访问的页面，默认读取当前日期的前一天的数据（有用户访问时，通过GETDATA.ASP读取远程INDEX55.ASP文件，如果没有保存过数据则保存）
'远程主机129上的INDEX55文件，先判断查询的日期的DBF文件是否生成，如果存在则生成带格式的文本输出


'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="conn.asp"-->
<!--#include file="GETDATA.asp"-->
<!--#include file="../inc/session.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'on error resume next
Dwt.Out "<html>"& vbCrLf
Dwt.Out "<head>" & vbCrLf
Dwt.Out "<title>信息管理系统轨道衡报表页面</title>"& vbCrLf
Dwt.Out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.Out "<link href='/css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out "<link href='/css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out"<script language=javascript src='/js/popselectdate.js'></script>"
Dwt.Out "</head>"& vbCrLf
Dwt.Out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'   style=""overflow: auto;"">"& vbCrLf
action=request("action")
select case action
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
	call main
 case "del"
    call del
end select	
sub del()
	dateinput=request("year")&"-"&request("month")&"-"&request("day")
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from qch where day=#"&dateinput&"#"
	rsdel.open sqldel,connjlhs,1,3
	'dwt.out "<Script Language=Javascript>history.go(-1);<Script>"
	'set rsdel=nothing
	sqldel="delete * from issave where day1=#"&dateinput&"#"
	rsdel.open sqldel,connjlhs,1,3
	set rsdel=nothing  
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
  
end sub


Sub main()
	'dim dateinput 
	dateinput=request("year")&"-"&request("month")&"-"&request("day")
	if isnull(replace(dateinput,"-","")) or replace(dateinput,"-","")="" then dateinput=DATE () - 1
    dateinput=CDate(dateinput)
'response.Write dateinput
	call getdatA(dateinput)
	sql="SELECT * from qch where day=#"&dateinput&"#"
if request("wupin")<>"" then sql=sql&" and  wupin like '%" &request("wupin")& "%' "	
	
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>轨道衡报表 "&dateinput&"</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf

	if getyear="" then getyear=year(dateinput)
	if getmonth="" then getmonth=month(dateinput)
	if getday="" then getday=day(dateinput)
	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
	dwt.out "		 <form method='post'  action='gdh.asp'  name='form' >"
	
	'response.Write dateinput
	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput-2)&"&month="&month(dateinput-2)&"&day="&day(dateinput-2)&"'>"&year(dateinput-2)&"年"&month(dateinput-2)&"月"&day(dateinput-2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput-1)&"&month="&month(dateinput-1)&"&day="&day(dateinput-1)&"'>"&year(dateinput-1)&"年"&month(dateinput-1)&"月"&day(dateinput-1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	dwt.out "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >	"
	dwt.out "	 <select name='year'></select>年<select name='month'></select>月<select name='day'></select>日 &nbsp;&nbsp;<input  type='submit' name='Submit' value=' 查看 ' style='cursor:hand;'>"
	dwt.out "		 <script type='text/javascript' src='/js/selectdate.js'></script>"


	if now()-dateinput>1 then 	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput+1)&"&month="&month(dateinput+1)&"&day="&day(dateinput+1)&"'>"&year(dateinput-1)&"年"&month(dateinput+1)&"月"&day(dateinput+1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-dateinput>2 then 	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput+2)&"&month="&month(dateinput+2)&"&day="&day(dateinput+2)&"'>"&year(dateinput+2)&"年"&month(dateinput+2)&"月"&day(dateinput+2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"


  sTime=dateinput   
  mTime=month(sTime)   
  dTime=day(sTime)  
  IF   mTime<10   THEN   
        mTime="0"&mTime   
  End   IF   
  IF   dTime<10   THEN   
        dTime="0"&dTime   
  End   IF  
  nowday=year(sTime)&mTime&dTime     '查询报表的天



	dwt.out "<a href='toexcel.asp?year="&getyear&"&month="&getmonth&"&day="&getday&"'>报表下载</a>"
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按种类跳转…</option>" & vbCrLf
	sqlgh="SELECT distinct wupin from qch where  day=#"&dateinput&"#"
	'if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	'if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    'sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connjlhs,1,1
    do while not rsgh.eof
		'cjid=cint(rsgh("sb_sscj"))
		'sql="SELECT count(sb_id) FROM sb WHERE sb_sscj="&cjid&"and  sb_dclass="&sb_classid
		'if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		'if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		'sb_numb=Conn.Execute(sql)(0)
        
		'if sb_numb<>0 then 
			'i=i+1
			Dwt.out"<option value='?year="&request("year")&"&month="&request("month")&"&day="&request("day")&"&wupin="&ltrim(rtrim(rsgh("wupin")))&"'"
			if request("wupin")=ltrim(rtrim(rsgh("wupin"))) then Dwt.out" selected"
			
			Dwt.out ">"&ltrim(rtrim(rsgh("wupin")))&"</option>"& vbCrLf '
	   ' end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf
	dwt.out "<a href='?action=del&year="&getyear&"&month="&getmonth&"&day="&getday&"'>刷新数据</a></form>	</div></div>"
															
			pz_total=0
			mz_total=0
			'pjs_total=cint(rs("pjs"))+pjs_total
			jz_total=0
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connjlhs,1,1
	if rs.eof and rs.bof then 
		Dwt.Out "<p align='center'>未添加内容</p>" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;' >"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""   style=""overflow: auto;"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header""   style=""overflow: auto;"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>序号</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>车号</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>单位</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>品种</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>速度</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>毛重</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>皮重</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>载重</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>净重</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>盈亏</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>过车时间</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		do while not rs.eof
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row'  onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
            if rs("yingkui")<0 then 
				xh_id1="<font color=red>"&xh_id&"</font>"
			else
			    xh_id1=xh_id
			end if	
			pz_total=Round(rs("pizhs"),3)+pz_total
			mz_total=Round(rs("maozhs"),3)+mz_total
			'pjs_total=cint(rs("pjs"))+pjs_total
			jz_total=Round(rs("jingzhs"),3)+jz_total
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id1&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("chehao")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("danwei")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("wupin")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("sudu")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("maozhs")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("pizhs")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("zaizhs")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("jingzhs")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("yingkui")&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&rs("gbdate")&"</Div></td>"& vbCrLf
			Dwt.Out "</tr>" & vbCrLf
	  rs.movenext
	  loop
			Dwt.Out "<tr class='x-grid-row ' bgcolor=#BFDFFF>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">合计</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id&"台</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""right"">毛重：</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&mz_total&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""right"">皮重：</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&pz_total&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""right"">净重：</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&jz_total&"</Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""right""></Div></td>"& vbCrLf
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center""></Div></td>"& vbCrLf
			Dwt.Out "</tr>" & vbCrLf

			
			
			Dwt.Out "</table>" & vbCrLf
		   Dwt.Out "</Div>"
		   end if
		   Dwt.Out "</Div>"		   
		   rs.close
		   set rs=nothing
end Sub
Dwt.Out "</body></html>"





Sub search()
end Sub




set connjlhs=nothing

Call Closeconn
%>