<%@language=vbscript codepage=936 %>
<%



'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
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
response.Write "<title>��Ϣ����ϵͳ�����ⱨ��ҳ��</title>"& vbCrLf
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
  nowday=year(sTime)&"-"&mTime&"-"&dTime     '��ѯ�������
'response.Write dateinput
'����
sql="SELECT * from [�����ܱ�] where [����ʱ��] like '%"&nowday&"%'"
if request("wupin")<>"" then sql=sql&" and  Ʒ�� like '%" &request("wupin")& "%' "
	'response.Write sql
	
	response.Write "<Div style='left:6px;'>"& vbCrLf
	response.Write "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	response.Write "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>�����ⱨ�� "&dateinput&"</span>"& vbCrLf
	response.Write "     </Div>"& vbCrLf

	if getyear="" then getyear=year(dateinput)
	if getmonth="" then getmonth=month(dateinput)
	if getday="" then getday=day(dateinput)
	response.Write "<div class='x-toolbar' style='padding-left:15px;'>"
	response.Write "	<div align=left>"
	response.Write "		 <form method='post'  action='bah.asp'  name='form' >"
	
	'response.Write dateinput
	response.Write "<a href='bah.asp?year="&year(dateinput-2)&"&month="&month(dateinput-2)&"&day="&day(dateinput-2)&"'>"&year(dateinput-2)&"��"&month(dateinput-2)&"��"&day(dateinput-2)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	response.Write "<a href='bah.asp?year="&year(dateinput-1)&"&month="&month(dateinput-1)&"&day="&day(dateinput-1)&"'>"&year(dateinput-1)&"��"&month(dateinput-1)&"��"&day(dateinput-1)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	response.Write "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >	"
	response.Write "	 <select name='year'></select>��<select name='month'></select>��<select name='day'></select>�� &nbsp;&nbsp;<input  type='submit' name='Submit' value=' �鿴 ' style='cursor:hand;'>"
	response.Write "		 <script type='text/javascript' src='/js/selectdate.js'></script>"


	if now()-dateinput>1 then 	response.Write "<a href='bah.asp?year="&year(dateinput+1)&"&month="&month(dateinput+1)&"&day="&day(dateinput+1)&"'>"&year(dateinput-1)&"��"&month(dateinput+1)&"��"&day(dateinput+1)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-dateinput>2 then 	response.Write "<a href='bah.asp?year="&year(dateinput+2)&"&month="&month(dateinput+2)&"&day="&day(dateinput+2)&"'>"&year(dateinput+2)&"��"&month(dateinput+2)&"��"&day(dateinput+2)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"


 
 
Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>��������ת��</option>" & vbCrLf
	sqlgh="SELECT distinct Ʒ�� from [�����ܱ�] where [����ʱ��] like '%"&nowday&"%'"
	
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,connjlhs,1,1
    do while not rsgh.eof
		
			Dwt.out"<option value='?year="&request("year")&"&month="&request("month")&"&day="&request("day")&"&wupin="&ltrim(rtrim(rsgh("Ʒ��")))&"'"
			if request("Ʒ��")=ltrim(rtrim(rsgh("Ʒ��"))) then Dwt.out" selected"
			
			Dwt.out ">"&ltrim(rtrim(rsgh("Ʒ��")))&"</option>"& vbCrLf '
	   
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf





	'response.Write "<a href=http://172.16.10.129/"&getyear&"/"&nowday&".xls>��������</form>	</div>"
	response.Write "</div></div>"
															

	set rs=server.createobject("adodb.recordset")
	'response.Write "<br>"&sql
	rs.open sql,connjlhs,1,1
	if rs.eof and rs.bof then 
		response.Write "<p align='center'>δ�������</p>" 
	else
		response.Write "<Div class='x-layOut-panel' style='WIDTH: 100%;' >"& vbCrLf
		
		response.Write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""   style=""overflow: auto;"">"& vbCrLf
		response.Write "<tr class=""x-grid-header""   style=""overflow: auto;"">" & vbCrLf
		response.Write "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>������λ</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>�ջ���λ</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>ë��</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>Ƥ��</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>����˾��Ա</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>ë������ </Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>����˾��Ա</Div></td>" & vbCrLf
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>Ƥ������</Div></td>" & vbCrLf	
		response.Write "<td class='x-td'><Div class='x-grid-hd-text'>ӯ��</Div></td>" & vbCrLf	
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
			if rs("Ƥ��")<>"" or not isnull(rs("Ƥ��")) then pz_total=CLng(rs("Ƥ��"))+pz_total
			if rs("ë��")<>"" or not isnull(rs("ë��")) then mz_total=CLng(rs("ë��"))+mz_total
			'pjs_total=cint(rs("pjs"))+pjs_total
			if rs("����")<>"" or not isnull(rs("����")) then jz_total=CLng(rs("����"))+jz_total
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("���")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("Ʒ��")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("����")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("��λ")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("����")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("ë��")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("Ƥ��")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("����")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("��������Ա")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("����ʱ��")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("��������Ա")&"</Div></td>"& vbCrLf
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("����ʱ��")&"</Div></td>"& vbCrLf	
		response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&rs("ӯ��")&"</Div></td>"& vbCrLf	

			response.Write "</tr>" & vbCrLf
	  rs.movenext
	  loop
			response.Write "<tr class='x-grid-row ' bgcolor=#BFDFFF>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">�ϼ�</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id&"̨</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">ë�أ�</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&mz_total&"</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">Ƥ�أ�</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""center"">"&pz_total&"</Div></td>"& vbCrLf
			response.Write "     <td  CLASS='X-TD'><Div align=""right"">���أ�</Div></td>"& vbCrLf
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