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
<!--#include file="conn.asp"-->
<!--#include file="../inc/function.asp"--><%

'on error resume next
response.Write "<html>"& vbCrLf
response.Write "<head>" & vbCrLf
response.Write "<title>信息管理系统苯胺衡报表页面</title>"& vbCrLf
response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.Write "<link href='/css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.Write "<link href='/css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.Write"<script language=javascript src='/js/popselectdate.js'></script>"
response.Write "</head>"& vbCrLf
response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'   style=""overflow: auto;"">"& vbCrLf
action=request("action")
select case action
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
	call main
	
end select	



Sub main()
	'dim dateinput 
	dateinput=request("year")&"-"&request("month")&"-"&request("day")
	if isnull(replace(dateinput,"-","")) or replace(dateinput,"-","")="" then dateinput=DATE () - 1
    dateinput=CDate(dateinput)

 sTime=dateinput   
  mTime=month(sTime)   
  dTime=day(sTime)  
  IF   mTime<10   THEN   
        mTime="0"&mTime   
  End   IF   
  IF   dTime<10   THEN   
        dTime="0"&dTime   
  End   IF  
  nowday=year(sTime)&"-"&mTime&"-"&dTime     '查询报表的天
'response.Write dateinput
'出货
sql="SELECT * from [数据总表] where [出厂时间] like '%"&nowday&"%'"
if request("wupin")<>"" then sql=sql&" and  品种 like '%" &request("wupin")& "%' "
	'response.Write sql
	
	response.Write "<Div style='left:6px;'>"& vbCrLf
	response.Write "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	response.Write "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>苯胺衡报表 "&dateinput&"</span>"& vbCrLf
	response.Write "     </Div>"& vbCrLf

	if getyear="" then getyear=year(dateinput)
	if getmonth="" then getmonth=month(dateinput)
	if getday="" then getday=day(dateinput)
	response.Write "<div class='x-toolbar' style='padding-left:15px;'>"
	response.Write "	<div align=left>"
	response.Write "		 <form method='post'  action='bah.asp'  name='form' >"
	
	'response.Write dateinput
	response.Write "<a href='bah.asp?year="&year(dateinput-2)&"&month="&month(dateinput-2)&"&day="&day(dateinput-2)&"'>"&year(dateinput-2)&"年"&month(dateinput-2)&"月"&day(dateinput-2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	response.Write "<a href='bah.asp?year="&year(dateinput-1)&"&month="&month(dateinput-1)&"&day="&day(dateinput-1)&"'>"&year(dateinput-1)&"年"&month(dateinput-1)&"月"&day(dateinput-1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	response.Write "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >	"
	response.Write "	 <select name='year'></select>年<select name='month'></select>月<select name='day'></select>日 &nbsp;&nbsp;<input  type='submit' name='Submit' value=' 查看 ' style='cursor:hand;'>"
	response.Write "		 <script type='text/javascript' src='/js/selectdate.js'></script>"


	if now()-dateinput>1 then 	response.Write "<a href='bah.asp?year="&year(dateinput+1)&"&month="&month(dateinput+1)&"&day="&day(dateinput+1)&"'>"&year(dateinput-1)&"年"&month(dateinput+1)&"月"&day(dateinput+1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-dateinput>2 then 	response.Write "<a href='bah.asp?year="&year(dateinput+2)&"&month="&month(dateinput+2)&"&day="&day(dateinput+2)&"'>"&year(dateinput+2)&"年"&month(dateinput+2)&"月"&day(dateinput+2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"


 
 
Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>按种类跳转…</option>" & vbCrLf
	sqlgh="SELECT distinct 品种 from [数据总表] where [出厂时间] like '%"&nowday&"%'"
	
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connjlhs,1,1
    do while not rsgh.eof
		
			Dwt.out"<option value='?year="&request("year")&"&month="&request("month")&"&day="&request("day")&"&wupin="&ltrim(rtrim(rsgh("品种")))&"'"
			if request("品种")=ltrim(rtrim(rsgh("品种"))) then Dwt.out" selected"
			
			Dwt.out ">"&ltrim(rtrim(rsgh("品种")))&"</option>"& vbCrLf '
	   
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf





	'response.Write "<a href=http://172.16.10.129/"&getyear&"/"&nowday&".xls>报表下载</form>	</div>"
	response.Write "</div></div>"
															

	set rs=server.createobject("adodb.recordset")
	'response.Write "<br>"&sql
	rs.open sql,connjlhs,1,1
	if rs.eof and rs.bof then 
		response.Write "<p align='center'>未添加内容</p>" 
	else
		response.Write "<Div class='x-layOut-panel' style='WIDTH: 100%;' >"& vbCrLf
		
		response.Write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""   style=""overflow: auto;"">"& vbCrLf
		response.Write "<tr class=""x-grid-header""   style=""overflow: auto;"">" & vbCrLf
		response.Write "     <td  class='x-td'><Div class='x-grid-hd-text'>序号</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>货物名称</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>车号</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>发货单位</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>收货单位</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>毛重</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>皮重</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>净重</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>进厂司磅员</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>毛重日期 </Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>出厂司磅员</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>皮重日期</Div></td>" & vbCrLf	
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>盈亏</Div></td>" & vbCrLf	
		response.Write "    </tr>" & vbCrLf
	

	
	
	
		do while not rs.eof
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  response.Write "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  response.Write "<tr class='x-grid-row'  onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
'            if rs("pjs")-rs("jz")>0.12 then 
'				xh_id1="<font color=red>"&xh_id&"</font>"
'			else
			    xh_id1=xh_id
'			end if	
			if rs("皮重")<>"" or not isnull(rs("皮重")) then pz_total=CLng(rs("皮重"))+pz_total
			if rs("毛重")<>"" or not isnull(rs("毛重")) then mz_total=CLng(rs("毛重"))+mz_total
			'pjs_total=cint(rs("pjs"))+pjs_total
			if rs("净重")<>"" or not isnull(rs("净重")) then jz_total=CLng(rs("净重"))+jz_total
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("序号")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("品种")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("车号")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("单位")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("车型")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("毛重")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("皮重")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("净重")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("进厂计量员")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("日期时间")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("出厂计量员")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("出厂时间")&"</Div></td>"& vbCrLf	
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("盈亏")&"</Div></td>"& vbCrLf	

			response.Write "</tr>" & vbCrLf
	  rs.movenext
	  loop
			response.Write "<tr class='x-grid-row ' bgcolor=#BFDFFF>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">合计</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id&"台</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">毛重：</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&mz_total&"</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">皮重：</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&pz_total&"</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">净重：</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&jz_total&"</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right""></Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center""></Div></td>"& vbCrLf
			response.Write "</tr>" & vbCrLf

			
			
			response.Write "</table>" & vbCrLf
		   response.Write "</Div>"
		   end if
		   response.Write "</Div>"		   
		   rs.close
		   set rs=nothing
end Sub
response.Write "</body></html>"









set connjlhs=nothing

%>