<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="GETDATA.asp"-->
<!--#include file="../inc/session.asp"-->
<!--#include file="../inc/function.asp"-->
<%
	
	
	
	Response.Buffer = True 
	Response.ContentType = "application/vnd.ms-excel" 
	Response.AddHeader "content-disposition", "inline; filename =轨道衡报表 "&request("year")&"-"&request("month")&"-"&request("day")&".xls"' 
'Response.AddHeader "content-disposition", "inline; filename =11111.xls"' 
	dateinput=request("year")&"-"&request("month")&"-"&request("day")
	if isnull(replace(dateinput,"-","")) or replace(dateinput,"-","")="" then dateinput=DATE () - 1
    dateinput=CDate(dateinput)

	sql="SELECT * from qch where day=#"&dateinput&"#"
	

'	if getyear="" then getyear=year(dateinput)
'	if getmonth="" then getmonth=month(dateinput)
'	if getday="" then getday=day(dateinput)
'	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
'	dwt.out "	<div align=left>"
'	dwt.out "		 <form method='post'  action='gdh.asp'  name='form' >"
'	
'	'response.Write dateinput
'	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput-2)&"&month="&month(dateinput-2)&"&day="&day(dateinput-2)&"'>"&year(dateinput-2)&"年"&month(dateinput-2)&"月"&day(dateinput-2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
'	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput-1)&"&month="&month(dateinput-1)&"&day="&day(dateinput-1)&"'>"&year(dateinput-1)&"年"&month(dateinput-1)&"月"&day(dateinput-1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
'	
'	dwt.out "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >	"
'	dwt.out "	 <select name='year'></select>年<select name='month'></select>月<select name='day'></select>日 &nbsp;&nbsp;<input  type='submit' name='Submit' value=' 查看 ' style='cursor:hand;'>"
'	dwt.out "		 <script type='text/javascript' src='/js/selectdate.js'><script>"
'
'
'	if now()-dateinput>1 then 	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput+1)&"&month="&month(dateinput+1)&"&day="&day(dateinput+1)&"'>"&year(dateinput-1)&"年"&month(dateinput+1)&"月"&day(dateinput+1)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
'	if now()-dateinput>2 then 	dwt.out "<a href='/gdh/gdh.asp?year="&year(dateinput+2)&"&month="&month(dateinput+2)&"&day="&day(dateinput+2)&"'>"&year(dateinput+2)&"年"&month(dateinput+2)&"月"&day(dateinput+2)&"日</a>&nbsp;&nbsp;&nbsp;&nbsp;"
'
'
'  sTime=dateinput   
'  mTime=month(sTime)   
'  dTime=day(sTime)  
'  IF   mTime<10   THEN   
'        mTime="0"&mTime   
'  End   IF   
'  IF   dTime<10   THEN   
'        dTime="0"&dTime   
'  End   IF  
'  nowday=year(sTime)&mTime&dTime     '查询报表的天
'
'
'
'	'dwt.out "<a href=http://172.16.10.129/"&getyear&"/"&nowday&".xls>报表下载"
'	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
'	Dwt.out "<option value=''>按种类跳转…</option>" & vbCrLf
'	sqlgh="SELECT distinct wupin from qch where  day=#"&dateinput&"#"
'	'if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
'	'if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
'    'sqlgh=sqlgh&" order by sb_sscj asc"
'	set rsgh=server.createobject("adodb.recordset")
'    rsgh.open sqlgh,connjlhs,1,1
'    do while not rsgh.eof
'		'cjid=cint(rsgh("sb_sscj"))
'		'sql="SELECT count(sb_id) FROM sb WHERE sb_sscj="&cjid&"and  sb_dclass="&sb_classid
'		'if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
'		'if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
'        
'		'sb_numb=Conn.Execute(sql)(0)
'        
'		'if sb_numb<>0 then 
'			'i=i+1
'			Dwt.out"<option value='?year="&request("year")&"&month="&request("month")&"&day="&request("day")&"&wupin="&ltrim(rtrim(rsgh("wupin")))&"'"
'			if request("wupin")=ltrim(rtrim(rsgh("wupin"))) then Dwt.out" selected"
'			
'			Dwt.out ">"&ltrim(rtrim(rsgh("wupin")))&"</option>"& vbCrLf '
'	   ' end if 
'		rsgh.movenext
'	loop
'	rsgh.close
'	set rsgh=nothing
'	Dwt.out "     </select>	" & vbCrLf
'	dwt.out "</form>	</div></div>"
															
			pz_total=0
			mz_total=0
			'pjs_total=cint(rs("pjs"))+pjs_total
			jz_total=0
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connjlhs,1,1
	if rs.eof and rs.bof then 
		'Dwt.Out "<p align='center'>未添加内容</p>" 
	else
		'Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;' >"& vbCrLf
		
		Dwt.Out "<table >"& vbCrLf
		Dwt.Out "<tr >" & vbCrLf
		Dwt.Out "     <td>序号</td>" & vbCrLf
		Dwt.Out "      <td>车号</td>" & vbCrLf
		Dwt.Out "      <td>单位</td>" & vbCrLf
		Dwt.Out "      <td>品种</td>" & vbCrLf
		Dwt.Out "      <td>速度</td>" & vbCrLf
		Dwt.Out "      <td>毛重</td>" & vbCrLf
		Dwt.Out "      <td>皮重</td>" & vbCrLf
		Dwt.Out "      <td>载重</td>" & vbCrLf
		Dwt.Out "      <td>净重</td>" & vbCrLf
		Dwt.Out "      <td>盈亏</td>" & vbCrLf
		Dwt.Out "      <td>过车时间</td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		do while not rs.eof
			dim xh,xh_id
			xh_id=1+xh_id
			  Dwt.Out "<tr >"& vbCrLf
			
			pz_total=clng(rs("pizhs"))+pz_total
			mz_total=clng(rs("maozhs"))+mz_total
			'pjs_total=cint(rs("pjs"))+pjs_total
			jz_total=clng(rs("jingzhs"))+jz_total
			Dwt.Out "     <td >"&xh_id&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("chehao")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("danwei")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("wupin")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("sudu")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("maozhs")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("pizhs")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("zaizhs")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("jingzhs")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("yingkui")&"</td>"& vbCrLf
			Dwt.Out "     <td>"&rs("gbdate")&"</td>"& vbCrLf
			Dwt.Out "</tr>" & vbCrLf
	  rs.movenext
	  loop
			Dwt.Out "<tr >"& vbCrLf
			Dwt.Out "     <td>合计</td>"& vbCrLf
			Dwt.Out "     <td>"&xh_id&"台</td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td>毛重：</td>"& vbCrLf
			Dwt.Out "     <td>"&mz_total&"</td>"& vbCrLf
			Dwt.Out "     <td>皮重：</td>"& vbCrLf
			Dwt.Out "     <td>"&pz_total&"</td>"& vbCrLf
			Dwt.Out "     <td>净重：</td>"& vbCrLf
			Dwt.Out "     <td>"&jz_total&"</td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "     <td></td>"& vbCrLf
			Dwt.Out "</tr>" & vbCrLf

			
			
			Dwt.Out "</table>" & vbCrLf
		   end if
		   rs.close
		   set rs=nothing

set connjlhs=nothing

Call Closeconn
%>