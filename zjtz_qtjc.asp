<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'dim starttime : starttime=timer
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<%
dim ydate
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh"))
 acdate=trim(request("acdate"))
sb_classid = Trim(Request("sbclassid"))
   if sb_classid="" then sb_classid=164
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sb_classid)(0)

Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>�ܼ�̨�ʹ���ҳ</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/grid.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/docs.css' rel='stylesheet' type='text/css'>"& vbCrLf

Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
   case "zjpost"
     call zjpost
   case "zjpost2"
     call zjpost2
   case "yzj"
     call yzj
   case "yzj2"
     call yzj2
  case "editinfo"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editinfo
  case "saveeditinfo"
    call saveeditinfo
  case "delinfo"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call delinfo

  case "history"
      call history
	

   case "addzjinfo"
     call addzjinfo
	 
   case "savezjinfo"
     call savezjinfo
	  
  case "del"
        if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 	
	   conn.Execute("UPDATE sbqt SET sb_iszj=false WHERE sb_id="&request("id"))
	  Dwt.out"<Script Language=Javascript>history.back();</Script>"
		end if 
  case ""
      if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	  	 

Sub zjpost()
	dim zjmonth
	zjyear=cint(request("zjyear"))
	zjmonth=cint(request("zjmonth"))
    sscj=request("sscj")
	ssbz=request("ssbz")
		
	url="zjtz_qtjc.asp?action=zjpost&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj&"&ssbz="&ssbz
	
	zjmonth_d=zjmonth&"��"
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&zjyear&"��-"&zjmonth_d&" "&sscjh(sscj)&" "&ssbzh(ssbz)&" ���ڼ���̨��</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf
	
if zjmonth<>0 then sql="SELECT * from sbqt where (year(dateadd('m',sb_test_period,sb_sczjdate))="&zjyear&" or year(sb_sczjdate)="&zjyear&") and (month(dateadd('m',sb_test_period,sb_sczjdate))="&zjmonth&"  or month(sb_sczjdate)="&zjmonth&") and sb_sscj="&sscj&" and sb_ssbz="&ssbz&" and sb_dclass=164 and sb_test_period<>0 and sb_iszj=true ORDER BY sb_id aSC "
	if zjmonth=0 then sql="SELECT * from sbqt where sb_sscj="&sscj&" and sb_ssbz="&ssbz&" and sb_dclass=164 and sb_iszj=true and sb_test_period<>0 ORDER BY sb_id aSC "

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		message "δ�ҵ��������" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>װ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ͺ�</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ƻ���������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ʵ�ʼ�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ƻ��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
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
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.Out "     <td  Class='x-td'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
						  
			if zclassor(rs("sb_dclass")) then 
			   if zclass(rs("sb_zclass"))="δ�༭" then 
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_dclass"))&"&nbsp;</td>" & vbCrLf
			   else
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_zclass"))&"&nbsp;</td>" & vbCrLf
			   end if 
			 end if  
			  	
			Dwt.Out "<td  class='x-td'>"&rs("sb_wh")&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&GH(rs("sb_ssGH"))&"&nbsp;</td>" & vbCrLf

			Dwt.Out "<td  class='x-td'>"&rs("sb_ggxh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_c1")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_jddj")&"&nbsp;</td>" & vbCrLf
			
			
			
			
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_test_period"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf

			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_jczq"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf


					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					dim jczq
					jdzq=rs("sb_test_period")/12
					jczq=rs("sb_jczq")/12


			Dwt.Out " <td  class='x-td' style='color:red'><Div align=""center"">"				   
			   Dwt.out  dateadd("m",rs("sb_test_period"),rs("sb_sczjdate"))
			Dwt.out "</Div></td>" & vbCrLf

			Dwt.Out " <td  class='x-td' style='color:red'><Div align=""center"">"				   
			   Dwt.out  rs("sb_sczjdate")
			Dwt.out "</Div></td>" & vbCrLf


			Dwt.Out " <td  class='x-td'><Div align=""center"">"				   
			   Dwt.out  dateadd("m",rs("sb_jczq"),rs("sb_scjcdate"))
			Dwt.out "</Div></td>" & vbCrLf
	
			dim sqlinfo,rsinfo
			dim c_text
			Dwt.Out "<td  class='x-td'><Div align=""center"">"

			
			if zjmonth<>0  then sqlinfo="SELECT * from zjinfo_qtjc where  year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("sb_id")
'			if zjmonth=0  then sqlinfo="SELECT * from zjinfo_qtjc where  dxzjyear="&zjyear&"  and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "δ�ܼ�"
					c_text="<a href=zjtz_qtjc.asp?action=addzjinfo&numb=1&id="&rs("sb_id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjdate="&zjyear&"-"&zjmonth&">���</a>  "

			    c_text=c_text&"  <a href=zjtz_qtjc.asp?action=addzjinfo&numb=1&id="&rs("sb_id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&">���ļƻ�����</a>"
			else
				 DWT.OUT RSINFO("zjdate")
				dim jdjg
				if rsinfo("zjinfo")="" then
				   jdjg="δ��д�������"
				else
				   jdjg=rsinfo("zjinfo")
				end if       
				c_text="�ܼ���� "&jdjg
			end if 
			
			Dwt.out "</Div></td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
			dwt.out c_text
			Dwt.Out "</Div></td></tr>" & vbCrLf
			c_text=""
			 RowCount=RowCount-1
	  rs.movenext
	  loop
	Dwt.Out "</table>" & vbCrLf
	   call showpage(page,url,total,record,PgSz)
   Dwt.Out "</Div>"
   end if
   Dwt.Out "</Div>"		   
   rs.close
   set rs=nothing
End Sub

'���ڱ��汾���ܼ���ɺ�������ܼ���


Sub zjpost2()
	dim zjmonth
	zjyear=cint(request("zjyear"))
	zjmonth=cint(request("zjmonth"))
    sscj=request("sscj")
	ssbz=request("ssbz")
		
	url="zjtz_qtjc.asp?action=zjpost2&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj&"&ssbz="&ssbz
	
	zjmonth_d=zjmonth&"��"
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&zjyear&"��-"&zjmonth_d&" "&sscjh(sscj)&" "&ssbzh(ssbz)&" �ճ����̨��</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf
	
if zjmonth<>0 then sql="SELECT * from sbqt where (year(dateadd('m',sb_jczq,sb_scjcdate))="&zjyear&" or year(sb_scjcdate)="&zjyear&") and (month(dateadd('m',sb_jczq,sb_scjcdate))="&zjmonth&"  or month(sb_scjcdate)="&zjmonth&") and sb_sscj="&sscj&" and sb_ssbz="&ssbz&" and sb_dclass=164 and sb_jczq<>0 and sb_iszj=true ORDER BY sb_id aSC "
	if zjmonth=0 then sql="SELECT * from sbqt where sb_sscj="&sscj&" and sb_ssbz="&ssbz&" and sb_dclass=164 and sb_iszj=true and sb_jczq<>0 ORDER BY sb_id aSC "

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		message "δ�ҵ��������" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>װ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ͺ�</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ƻ��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ʵ�ʼ������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ƻ���������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
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
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.Out "     <td  Class='x-td'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
						  
			if zclassor(rs("sb_dclass")) then 
			   if zclass(rs("sb_zclass"))="δ�༭" then 
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_dclass"))&"&nbsp;</td>" & vbCrLf
			   else
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_zclass"))&"&nbsp;</td>" & vbCrLf
			   end if 
			 end if  
			  	
			Dwt.Out "<td  class='x-td'>"&rs("sb_wh")&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&GH(rs("sb_ssGH"))&"&nbsp;</td>" & vbCrLf

			Dwt.Out "<td  class='x-td'>"&rs("sb_ggxh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_c1")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_jddj")&"&nbsp;</td>" & vbCrLf
			
			
			
			
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_test_period"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf

			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_jczq"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf


					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					dim jczq
					jdzq=rs("sb_test_period")/12
					jczq=rs("sb_jczq")/12
					
			Dwt.Out " <td  class='x-td' style='color:red'><Div align=""center"">"				   
			   Dwt.out  dateadd("m",rs("sb_jczq"),rs("sb_scjcdate"))
			Dwt.out "</Div></td>" & vbCrLf
			
			Dwt.Out " <td  class='x-td' style='color:red'><Div align=""center"">"				   
			   Dwt.out  rs("sb_scjcdate")
			Dwt.out "</Div></td>" & vbCrLf

			Dwt.Out " <td  class='x-td' ><Div align=""center"">"				   
			   Dwt.out  dateadd("m",rs("sb_test_period"),rs("sb_sczjdate"))
			Dwt.out "</Div></td>" & vbCrLf


	
			dim sqlinfo,rsinfo
			dim c_text
			Dwt.Out "<td  class='x-td'><Div align=""center"">"

			
			if zjmonth<>0  then sqlinfo="SELECT * from zjinfo_qtjc where  year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("sb_id")
'			if zjmonth=0  then sqlinfo="SELECT * from zjinfo_qtjc where  dxzjyear="&zjyear&"  and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "δ�ܼ�"
					c_text="<a href=zjtz_qtjc.asp?action=addzjinfo&numb=0&id="&rs("sb_id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjdate="&zjyear&"-"&zjmonth&">���</a>  "

			    c_text=c_text&"  <a href=zjtz_qtjc.asp?action=addzjinfo&numb=0&id="&rs("sb_id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&">���ļƻ�����</a>"
			else
				 DWT.OUT RSINFO("zjdate")
				dim jdjg
				if rsinfo("zjinfo")="" then
				   jdjg="δ��д�������"
				else
				   jdjg=rsinfo("zjinfo")
				end if       
				c_text="�ܼ���� "&jdjg
			end if 
			
			Dwt.out "</Div></td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
			dwt.out c_text
			Dwt.Out "</Div></td></tr>" & vbCrLf
			c_text=""
			 RowCount=RowCount-1
	  rs.movenext
	  loop
	Dwt.Out "</table>" & vbCrLf
	   call showpage(page,url,total,record,PgSz)
   Dwt.Out "</Div>"
   end if
   Dwt.Out "</Div>"		   
   rs.close
   set rs=nothing
End Sub



sub yzj()
	Dwt.out "<br/><br/><br/><br/><br/>"
	dwt.out "<Div align='center'><Div class='x-dlg x-dlg-closable x-dlg-draggable x-dlg-modal' style=' WIDTH: 400px; HEIGHT: 198px'>"
	Dwt.out "  <Div class='x-dlg-hd-left'>"
	Dwt.out "    <Div class='x-dlg-hd-right'>"
	Dwt.out "      <Div class='x-dlg-hd x-unselectable'>�ܼ��豸��ѯ</Div>"
	Dwt.out "    </Div>"
	Dwt.out "  </Div>"
	Dwt.out "  <Div class='x-dlg-dlg-body' style='WIDTH: 400px;'><Div align=left>"

	Dwt.out"<br/><form method='post' action='zjtz_qtjc.asp' name='form1' onsubmit='javascript:return check();'>"
	Dwt.out "<table width='100%' >"& vbCrLf
	Dwt.out"<tr><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td> "
	Dwt.out"<td width='60%' class='tdbg'>"& vbCrLf
	Dwt.out "<select name='zjyear'>" & vbCrLf
	Dwt.out "<option value=''>ѡ�����</option>" & vbCrLf
	for i=year(now())-5 to year(now())+5
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=year(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
			'Dwt.out"<option value='"&i&"'>"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>��	" & vbCrLf
	Dwt.out "<select name='zjmonth'>" & vbCrLf
	Dwt.out "<option value=''>ѡ���·�</option>" & vbCrLf
	dwt.out "<option value=0>����</option>"
	for i=1 to 12
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=month(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>	" & vbCrLf
	Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	Dwt.out"<strong>�������䣺</strong></td>"
	Dwt.out "<td>" & vbCrLf
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	Dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'"& vbCrLf
		'if session("level")=rscj("levelid") then Dwt.out "selected"
		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
    Dwt.out "<select name='ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>ѡ��������</option>" & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
    Dwt.out "<script>" & vbCrLf
    Dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""�ް���"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out "var temp=document.form1.ssbz" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//</script" & vbCrLf
	Dwt.out "</td></tr>" & vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='zjpost'><input  type='submit' name='Submit' value='��ѯ' style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

	Dwt.out "  </Div></Div>"
	Dwt.out "</Div></Div>"
end sub


sub yzj2()
	Dwt.out "<br/><br/><br/><br/><br/>"
	dwt.out "<Div align='center'><Div class='x-dlg x-dlg-closable x-dlg-draggable x-dlg-modal' style=' WIDTH: 400px; HEIGHT: 198px'>"
	Dwt.out "  <Div class='x-dlg-hd-left'>"
	Dwt.out "    <Div class='x-dlg-hd-right'>"
	Dwt.out "      <Div class='x-dlg-hd x-unselectable'>�ܼ��豸��ѯ</Div>"
	Dwt.out "    </Div>"
	Dwt.out "  </Div>"
	Dwt.out "  <Div class='x-dlg-dlg-body' style='WIDTH: 400px;'><Div align=left>"

	Dwt.out"<br/><form method='post' action='zjtz_qtjc.asp' name='form1' onsubmit='javascript:return check();'>"
	Dwt.out "<table width='100%' >"& vbCrLf
	Dwt.out"<tr><td width='20%' align='right' class='tdbg'><strong>�ܼ��·ݣ�</strong></td> "
	Dwt.out"<td width='60%' class='tdbg'>"& vbCrLf
	Dwt.out "<select name='zjyear'>" & vbCrLf
	Dwt.out "<option value=''>ѡ�����</option>" & vbCrLf
	for i=year(now())-5 to year(now())+5
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=year(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
			'Dwt.out"<option value='"&i&"'>"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>��	" & vbCrLf
	Dwt.out "<select name='zjmonth'>" & vbCrLf
	Dwt.out "<option value=''>ѡ���·�</option>" & vbCrLf
	dwt.out "<option value=0>����</option>"
	for i=1 to 12
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=month(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>	" & vbCrLf
	Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	Dwt.out"<strong>�������䣺</strong></td>"
	Dwt.out "<td>" & vbCrLf
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	Dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'"& vbCrLf
		'if session("level")=rscj("levelid") then Dwt.out "selected"
		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
    Dwt.out "<select name='ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>ѡ��������</option>" & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
    Dwt.out "<script>" & vbCrLf
    Dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""�ް���"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out "var temp=document.form1.ssbz" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//</script" & vbCrLf
	Dwt.out "</td></tr>" & vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='zjpost2'><input  type='submit' name='Submit' value='��ѯ' style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

	Dwt.out "  </Div></Div>"
	Dwt.out "</Div></Div>"
end sub





sub main()
 url= GetUrl
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function isDel(id){" & vbCrLf
Dwt.out "		if(confirm('��ȷ��Ҫɾ����������')){" & vbCrLf
Dwt.out "			location.href='zjtz_qtjc.asp?action=del&id='+id;" & vbCrLf
Dwt.out "		}" & vbCrLf
Dwt.out "	}" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf


	if request("sscj")<>"" then title=sscjh(sscjid)&"��" 
	if request("ssgh")<>"" then title=gh(ssghid) &"��"
	if request("keyword")<>"" then title=" '"&keys&" '"&"��"
    title="����"&title&sb_classname
	if request("sbzclassid")<>"" then title=title&"��"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&request("sbzclassid"))(0)
	
	
	Dwt.out "<Div style='left:6px;'>"
	Dwt.out "     <Div class='x-layout-panel-hd x-layout-title-center'>"
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�ܼ�̨��"&title&"</span>"
	Dwt.out "     </Div>"

        sqlcj="SELECT distinct sb_sscj from sbqt where  sb_isdel=false and sb_iszj=true and sb_dclass="&sb_classid 
		
		   sqlcj=sqlcj&" order by sb_sscj asc"
	   set rscj=server.createobject("adodb.recordset")
               rscj.open sqlcj,conn,1,1
               do while not rscj.eof
	       sscji=cint(rscj("sb_sscj"))
           ' for sscji=1 to 5 
	  sql="SELECT count(sb_id) FROM sbqt WHERE sb_dclass="&sb_classid&" and sb_iszj=true and sb_sscj="&sscji
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	  sb_numb=sb_numb&sscjh_d(sscji)&":"&"<font color='#006600'>"&conn.Execute(sql)(0)&"</font>&nbsp;&nbsp;&nbsp;&nbsp;"
	   ' next
              rscj.movenext
	      loop
	      rscj.close
	      set rscj=nothing

	sql="SELECT count(sb_id) FROM sbqt WHERE sb_iszj=true and sb_dclass="&sb_classid
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	totall= "<font color='#006600'>"&conn.Execute(sql)(0)&"</font>" 
	Dwt.out "<Div class='pre'>"&sb_numb&"�ϼ�:"&totall&"<br/></Div>"& vbCrLf
	Dwt.out "<Div class='x-layout-container' style='top:0px;WIDTH: 100%; POSITION: relative; HEIGHT: 543px'>"& vbCrLf
	Dwt.out "<Div class='x-layout-panel x-layout-panel-center' style='LEFT: 3px; WIDTH: 100%; TOP: 3px; HEIGHT: 537px'>"& vbCrLf
	search	()
	
	
	sqlbody="SELECT * from sbqt where sb_isdel=false and sb_iszj=true and sb_dclass="&sb_classid
	if sscjid<>"" then sqlbody=sqlbody&" and sb_sscj="&sscjid
	if ssghid<>"" then sqlbody=sqlbody&" and sb_ssgh="&ssghid
	if keys<>"" then sqlbody=sqlbody&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlbody=sqlbody&" and sb_zclass="&request("sbzclassid")
	if request("update")<>"" then 
    	sqlbody=sqlbody&" order by sb_update desc"
	else
    	sqlbody=sqlbody&" order by sb_sscj aSC,sb_scjcdate asc,sb_wh asc"
	end if 

	set rsbody=server.createobject("adodb.recordset")
	rsbody.open sqlbody,conn,1,1

	if rsbody.eof and rsbody.bof then 
		message "<p align=""center"">δ�ҵ��������</p>" & vbCrLf
	else
	
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>װ��</Div></td>" & vbCrLf
		
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ʽ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>����ͺ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�ϴμ���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>�´μ���</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
	
		
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
		
				
		do while not rsbody.eof and rowcount>0
				xh_id=((page-1)*pgsz)+1+xh
				xh=xh+1
				
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 

			  Dwt.Out "<td  CLASS='X-TD'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
			  
			
			  Dwt.Out "<td class='x-td' ><Div align=""center"">"&sscjh_d(rsbody("sb_sscj"))
                call edit2(rsbody("sb_id"),rsbody("sb_sscj"))
              DWT.OUT "</Div></td>" & vbCrLf
			  
			  
			if zclassor(rsbody("sb_dclass")) then 
			   if zclass(rsbody("sb_zclass"))="δ�༭" then 
			Dwt.Out "<td  class='x-td'>"&zclass(rsbody("sb_dclass"))&"&nbsp;</td>" & vbCrLf
			   else
			Dwt.Out "<td  class='x-td'>"&zclass(rsbody("sb_zclass"))&"&nbsp;</td>" & vbCrLf
			   end if 
			 end if   	
			
			Dwt.Out "<td  class='x-td'>"
			if now()-rsbody("sb_update")<7 then Dwt.out "<span style=""color:#0033ff"">��</span>"
			Dwt.Out searchH(uCase(rsbody("sb_wh")),keys)&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&GH(rsbody("sb_ssGH"))&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("����ʽ",rsbody("sb_glfs"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&rsbody("sb_bh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rsbody("sb_ggxh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rsbody("sb_jddj")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rsbody("sb_c1")&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rsbody("sb_test_period"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf
			
					dim jdzq  '�춨�����ж�
					dim jdinfo
					dim jdyear '�춨���ڻ���Ϊ��
					jdzq=rsbody("sb_test_period")/12

			Dwt.Out " <td  class='x-td'><Div align=""center"">"				   
               if  rsbody("sb_test_period")<>1 then				
			     Dwt.out rsbody("sb_sczjdate")
			   end if 	 	 
			Dwt.out "</Div></td>" & vbCrLf
			
			'�´��ܼ�����
			Dwt.Out "<td  class='x-td'><Div align=""center"">"
                if  rsbody("sb_test_period")<>1 then			
                 Dwt.out dateadd("m",rsbody("sb_test_period"),rsbody("sb_sczjdate"))
			    end if 	 	 
            Dwt.out "</Div></td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&rsbody("sb_bz")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'><Div align=center>" & vbCrLf
					'call edit(rsbody("sb_id"),rsbody("sb_sscj"))
			Dwt.Out "</Div></td></tr>" & vbCrLf
			
			
				
			RowCount=RowCount-1
		rsbody.movenext
		loop
		
		Dwt.Out "</table>" & vbCrLf
		
	Dwt.out "<TABLE cellSpacing=0 cellPadding=0 border=0>"
	Dwt.out "  <TBODY>"
	Dwt.out "  <tr>"
	Dwt.out "    <TD valign='top' style='BORDER-RIGHT: white 2px inset; BORDER-TOP: white 2px inset; BORDER-LEFT: white 2px inset; BORDER-BOTTOM: white 2px inset; BACKGROUND-COLOR: scrollbar'>"
	Dwt.out "      <Div id=DataTable></Div></TD></tr></TBODY></TABLE>"
		call sbshowpage(page,url,total,record)
		   Dwt.Out "</Div>"
		   end if
		   Dwt.Out "</Div>"		   
	rsbody.close
	set rsbody=nothing
	conn.close
	set conn=nothing

end sub

Dwt.out "</body></html>"

sub addzjinfo()
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from sbqt where sb_id="&request("id")&" ORDER BY sb_id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,conn,1,1
   if rszjtz.eof and rszjtz.bof then 
        message("δ֪����")
   else
	   Dwt.out"<br><br><br><form method='post' action='zjtz_qtjc.asp' name='form2' onsubmit='javascript:return addzjinfo();'>"
	   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
	   Dwt.out"<Div align='center'><strong>���������д</strong></Div></td>    </tr>"
	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"      
	   Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&sscjh(rszjtz("sb_sscj"))&"' size=10>&nbsp;</td></tr>"& vbCrLf
	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����װ�ã� </strong></td>"      
	   
	   Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&GH(rszjtz("sb_ssgh"))&"' size=10>&nbsp;</td></tr>"& vbCrLf
		
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
		Dwt.out"<strong>λ&nbsp;&nbsp;�ţ�</strong></td>"
		Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("sb_wh")&"></td>    </tr>   "
		 
		 
		 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;�ͣ�</strong></td> "
		 
	if zclassor(rszjtz("sb_dclass")) then
		Dwt.out"<td width='60%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass 164,rszjtz("sb_zclass") 
		Dwt.out"</select></td></tr>"& vbcrlf
    end if 
		 
		 
		 
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ͺţ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("sb_ggxh")&"></td>    </tr>   "
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������Χ��</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("sb_c1")&"></td>    </tr>   "
		 	if request("numb")=1 then	 
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input disabled='disabled' type='text' value="	
			dwt.out dispalydatadict("��������",rszjtz("sb_test_period"))
	dwt.out "></td></tr>"
	
			else
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input disabled='disabled' type='text' value="	
			dwt.out dispalydatadict("��������",rszjtz("sb_jczq"))
	dwt.out "></td></tr>"
	       end if
		
		 
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
    Dwt.out"<input name='zjtz_date' "
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&request("zjdate")&"'/>�ճ��ܼ�����"	
	
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ݣ�</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input name='zjinfo' type='text'></td>    </tr>   "
		
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���������</strong></td>"
		 Dwt.out"<td width='88%' class='tdbg'><input name='zjresult' type='text'></td>    </tr>   "
		 
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�� �� �ˣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_ren' type='text'><br>�����ֵ�����,�����м�������ӿո�������ַ� <br>���������,ÿ���˵������м����ÿո�����,����ʹ�������ַ�</td></tr>"& vbCrLf
	
		Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;ע��</strong></td>"      
		Dwt.out"<td width='88%' class='tdbg'><input type='text' name='bz'></td></tr>  "   
	
		Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
		Dwt.out"<input name='action' type='hidden' id='action' value='savezjinfo'> <input type='hidden' name='id' value='"&request("id")&"'> <input type='hidden' name='ssbz' value='"&rszjtz("sb_ssbz")&"'><input type='hidden' name='numb' value='"&request("numb")&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'></td>  </tr>"
		Dwt.out"</table></form>"
		'Dwt.out request("sscj")&&
   end if 
end sub


sub savezjinfo()
      dim rsadd,sqladd,temp_id
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zjinfo_qtjc" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("zjtzid")=Trim(Request("id"))
	     rsadd("zjdate")=request("zjtz_date")
		 zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
      rsadd("bz")=request("bz")
      rsadd("jx_ren")=request("jx_ren")
      rsadd("zjinfo")=request("zjinfo")
      rsadd("zjresult")=request("zjresult")
      rsadd("zj_numb")=request("zj_numb")
	  rsadd.update
	  temp_id=rsadd("id")
      rsadd.close
	 
	 
	  dim rsedit,sqledit
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbqt where sb_id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,conn,1,3
	  if request("numb")=1 then
	     rsedit("sb_sczjdate")=request("zjtz_date")
	  else
	     rsedit("sb_scjcdate")=request("zjtz_date")
		 end if
	     rsedit("sb_update")=now()
	  rsedit.update
      sscj=rsedit("sb_sscj")
      ssbz=rsedit("sb_ssbz")
      rsedit.close
      set rsedit=nothing
	  
	  dim rsjx,sqljx
	  
	set rsjx=server.createobject("adodb.recordset")
	sqljx="select * from sbjx" 
	rsjx.open sqljx,conn,1,3
	rsjx.addnew	  
	if request("numb")=1 then
	rsjx("jx_lb")=11
	else
	rsjx("jx_lb")=1
	end if
	rsjx("jx_gzxx")=33
	rsjx("jx_nr")=ReplaceBadChar(Trim(request("zjinfo")))
	rsjx("jx_gzxx_new")=33
	rsjx("jx_nr_new")=0
	rsjx("jx_date")=Trim(request("zjtz_date"))
	rsjx("jx_enddate")=Trim(request("zjtz_date"))
'	rsjx("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsjx("jx_ren")=ReplaceBadChar(Trim(request("jx_ren")))
	rsjx("jx_bz")=ReplaceBadChar(Trim(request("bz")))
	rsjx("jx_zjh")=temp_id
	rsjx("sb_id")=ReplaceBadChar(Trim(Request("id")))
	rsjx.update
	rsjx.close
      set rsjx=nothing
	  dim numb
	  numb=request("numb")
	if numb=1 then 
	  Dwt.out"<Script Language=Javascript>location.href='zjtz_qtjc.asp?action=zjpost&sscj="&sscj&"&ssbz="&ssbz&"&zjyear="&zjyear&"&zjmonth="&zjmonth&"';</Script>"
	  else
	  Dwt.out"<Script Language=Javascript>location.href='zjtz_qtjc.asp?action=zjpost2&sscj="&sscj&"&ssbz="&ssbz&"&zjyear="&zjyear&"&zjmonth="&zjmonth&"';</Script>"
	  end if
end sub

Sub editinfo()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zjinfo_qtjc where id="&id
   rsedit.open sqledit,connzj,1,1
   Dwt.Out"<br><br><br><form method='post' action='zjtz_qtjc.asp' name='form1' >"& vbCrLf
   Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   Dwt.Out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.Out"<Div align='center'><strong>�༭�ܼ���ʷ</strong></Div></td>    </tr>"& vbCrLf
	
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ڣ�</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
    Dwt.out"<input name='zjtz_date' "
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("zjdate")&"'/>�ճ��ܼ�����"	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ����ݣ�</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjinfo' type='text' value="&rsedit("zjinfo")&"></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�ܼ�����</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjresult' type='text' value="&rsedit("zjresult")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��ע��</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='bz' type='text' value="&rsedit("bz")&"></td>    </tr>   "
	
	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"& vbCrLf
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveeditinfo'> <input type='hidden' name='numb' value=1><input type='hidden' name='id' value='"&id&"'>      <input  type='Submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"& vbCrLf
	Dwt.Out"</table></form>"& vbCrLf
	       rsedit.close
       set rsedit=nothing
end Sub
sub saveeditinfo()
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from zjinfo_qtjc where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
	     rsedit("zjdate")=request("zjtz_date")
		 zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
      zjtzid=rsedit("zjtzid")
	  rsedit("bz")=request("bz")
	  rsedit("zjinfo")=request("zjinfo")
      rsedit("zjresult")=request("zjresult")
	  dim numb
	  numb=rsedit("zj_numb")
	  rsedit.update
      set rsedit=nothing
	  
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbqt where sb_id="&zjtzid
      rsedit.open sqledit,conn,1,3
	  if numb=1 then
	     rsedit("sb_sczjdate")=request("zjtz_date")
		 else
	     rsedit("sb_scjcdate")=request("zjtz_date")
	  end if
	     rsedit("sb_update")=now()
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  dim rsjx,sqljx
	  
	set rsjx=server.createobject("adodb.recordset")
	sqljx="select * from sbjx where jx_zjh="&ReplaceBadChar(Trim(request("id")))
	
	rsjx.open sqljx,conn,1,3
	if not rsjx.eof and not rsjx.bof then 
	rsjx("jx_nr")=ReplaceBadChar(Trim(request("zjinfo")))
	rsjx("jx_nr_new")=0
	rsjx("jx_date")=Trim(request("zjtz_date"))
	rsjx("jx_enddate")=Trim(request("zjtz_date"))
'	rsjx("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsjx("jx_zjh")=ReplaceBadChar(Trim(Request("id")))
	rsjx.update
	end if
	rsjx.close
      set rsjx=nothing
  Dwt.Out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


'ȡ�ӷ�������
function zclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclass= "δ�༭"
  else
     zclass=rsbody("sbclass_name")
  end if
end function

'�ж��Ƿ����ӷ���
function zclassor(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_zclass="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclassor=false 
  else
     zclassor=true
  end if
end function


'�������б���ʾ
function formdclass()
	dim sqldclass,rsdclass
	'if isnull(dclassid) then dclassid=0
'	if dclassid=0 then 
		sqldclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_isput=true"
'	else
'		sqldclass="SELECT * from sbclass where sbclass_dclass<>0 and sbclass_id="&dclassid
'	end if 		
	set rsdclass=server.createobject("adodb.recordset")
	rsdclass.open sqldclass,conn,1,1
	if rsdclass.eof then 
		dclass="û���κη���"
	else
		Dwt.out"<option value='0'"
		if dclassid=0 then Dwt.out " selected" 
			Dwt.out">��ѡ��Ҫ����豸�ķ���</option>"
		do while not rsdclass.eof
			Dwt.out"<option value='zjtz_qtjc.asp?action=addsb&sbclassid="&rsdclass("sbclass_id")&"'>"&rsdclass("sbclass_name")&"</option>"  & vbCrLf   
		rsdclass.movenext
		loop
	end if 
	rsdclass.close
	set rsdclass=nothing
end function


'�ӷ����б���ʾ
function formzclass(dclassid,zclassid)
	dim sqlzclass,rszclass
	if isnull(zclassid) then zclassid=0
'	if zclassid=0 then 
		sqlzclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_zclass="&dclassid
'	else
		'sqlzclass="SELECT * from sbclass where sbclass_zclass<>0 and sbclass_id="&zclassid
'	end if 		
	set rszclass=server.createobject("adodb.recordset")
	rszclass.open sqlzclass,conn,1,1
	if rszclass.eof then 
		formzclass="δ�༭"
	else
		Dwt.out"<option value='0'"
		if zclassid=0 then Dwt.out " selected" 
			Dwt.out">��ѡ������</option>"
		do while not rszclass.eof
			Dwt.out"<option value='"&rszclass("sbclass_id")&"' "
			if zclassid=rszclass("sbclass_id") then Dwt.out "selected"
			Dwt.out">"&rszclass("sbclass_name")&"</option>"  & vbCrLf   
		rszclass.movenext
		loop
	end if 
	rszclass.close
	set rszclass=nothing
end function

'��������ʾ
Function sb_whd(whnumb)
	if isnull(whnumb) or whnumb=0 then 
	  sb_whd="δ�༭"
	else
		if whnumb=1 then sb_whd="<span style=""color:#006600"">��</span>"  '�����
		if whnumb=2 then sb_whd="<span style=""color:#ff0000"">��</span> "	  '����ú�
	end if 
end Function 

'׼ȷ�����ʾ
Function sb_zqd(zqnumb)

	if isnull(zqnumb) or zqnumb=0 then 
	  sb_zqd="δ�༭"
	else
		if zqnumb=3 then sb_zqd="����"'>95%
		if zqnumb=2 then sb_zqd="���"		  '�м�  
		if zqnumb=1 then sb_zqd="��"  '�����С
	end if 
end Function 

'Ͷ�������ʾ
Function sb_tyd(tynumb)

	if isnull(tynumb) or tynumb=0 then 
	  sb_tyd="δ�༭"
	else
		if tynumb=1 then sb_tyd="<span style=""color:#006600"">��</span>"   '��Ͷ��
		if tynumb=2 then sb_tyd="<span style=""color:#0000ff"">��</span>"   '����ԭ��δͶ��
		if tynumb=3 then sb_tyd="<span style=""color:#ff0000"">��</span>"    '�칤��ԭ��δͶ��
		'if zqnumb=4 then sb_zqd="<font color='#ff0000'>*</font>"    '�칤��ԭ��δͶ��
	end if 
end Function 

Sub history()

    sql="SELECT * from sbqt where sb_id="&request("id")&" ORDER BY sb_id DESC"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    if rs.eof and rs.bof then 
        Dwt.Out "<p align='center'>δ�ҵ�����</p>" 
    else
		Dwt.Out "<Div style='left:6px;'>"& vbCrLf
		Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
		Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&rs("sb_wh")&"  �ܼ���ʷ</span>"& vbCrLf
		Dwt.Out "     </Div>"& vbCrLf
       
		'Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>λ��</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>װ��</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>����ʽ</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�ͺ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>���ȵȼ�</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>������Χ</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>��������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
	    Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sb_sscj"))&"</Div></td>" & vbCrLf
		
			if zclassor(rs("sb_dclass")) then 
			   if zclass(rs("sb_zclass"))="δ�༭" then 
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_dclass"))&"&nbsp;</td>" & vbCrLf
			   else
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_zclass"))&"&nbsp;</td>" & vbCrLf
			   end if 
			 end if   	

			Dwt.Out "<td  class='x-td'>"&searchH(uCase(rs("sb_wh")),keys)&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&GH(rs("sb_ssGH"))&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("����ʽ",rs("sb_glfs"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf
			
        Dwt.Out " <td  class='x-td'>"&rs("sb_sccj")&"&nbsp;</td>" & vbCrLf
		
			Dwt.Out "<td  class='x-td'>"&rs("sb_bh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_ggxh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_jddj")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_c1")&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_test_period"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf
	   
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("��������",rs("sb_jczq"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf
	   
        Dwt.Out "</tr></table>" & vbCrLf
	  sscjid=rs("sb_sscj")
	end if
	
	
    rs.close
    set rs=nothing
	
	sqlscdate="SELECT * from zjinfo_qtjc where zjtzid="&request("id")&" ORDER BY id DESC"
    set rsscdate=server.createobject("adodb.recordset")
    rsscdate.open sqlscdate,connzj,1,1
    if rsscdate.eof and rsscdate.bof then 
        message("û����ǰ���ܼ��¼")
    else
         record=rsscdate.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsscdate.PageSize = Cint(PgSz) 
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
           rsscdate.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsscdate.PageSize
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>���</Div></td>" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>�ܼ�����</Div></td>" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>�ܼ�����</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>��ע</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>ѡ��</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
		   do while not rsscdate.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&xh_id&"</Div></td>" & vbCrLf
        'zjmonth=month(rsscdate("zjdate"))
		'if zjmonth=0 then zjmonth="����"
                if rsscdate("isdx") then
                      zjdate=rsscdate("dxzjyear")&"-����"
                else
                      zjdate=rsscdate("zjdate")
                end if 
		Dwt.Out "      <td  class='x-td'>"&zjdate&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("zjinfo")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("zjresult")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("bz")&"&nbsp;</td>" & vbCrLf
		'dwt.out session("levelclass")&"-"&sscjid
		if session("levelclass")=sscjid or session("levelclass")=10 then 
			Dwt.Out "<td  class='x-td'><a href=zjtz_qtjc.asp?action=editinfo&id="&rsscdate("id")&">�༭</a>&nbsp;"
			Dwt.Out "<a href=zjtz_qtjc.asp?action=delinfo&id="&rsscdate("id")&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a></td>"
		 else
			Dwt.Out "&nbsp;"
		 end if 
 
			 RowCount=RowCount-1
          rsscdate.movenext
          loop
        Dwt.Out "</table>" & vbCrLf
       url="zjtz_qtjc.asp?action=history&id="&request("id")
	   call showpage(page,url,total,record,PgSz)
	   Dwt.Out "</Div>"
	   end if
	   Dwt.Out "</Div>"
	          rsscdate.close
	         Dwt.Out "<br><table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td>" 
			Dwt.Out "<input name='Cancel' type='button' id='Cancel' value=' ��  �� ' onClick="";history.back()"" style='cursor:hand;'></td></tr></table>"

end Sub

sub search()
	dim rscj,sqlcj,sscjid
	Dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	
	Dwt.out "<form method='Get' name='SearchForm' action='zjtz_qtjc.asp'>" & vbCrLf
	Dwt.out "&nbsp;&nbsp;<select   onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>��ʾ˳��ѡ��</option>" & vbCrLf
	Dwt.out "<option value='zjtz_qtjc.asp?update=update&sbclassid="&sb_classid&"'>������ʱ��</option>"
	Dwt.out "     </select>	" & vbCrLf

	
	Dwt.out "  <input type='hidden' name='sbclassid' value='"&sb_classid&"'>" & vbCrLf
	if request("sbzclassid")<>"" then Dwt.out "<input type='hidden' name='sbzclassid' value='"&request("sbzclassid")&"'>" & vbCrLf

	Dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 Dwt.out "value='"&request("keyword")&"'"
    	Dwt.out ">" & vbCrLf
    else
	 Dwt.out "value='����������λ��'"
	 	Dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	Dwt.out "  <input type='submit' name='Submit'  value='����'>"
	



	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tosscj(this.options[this.selectedIndex].value);}"">" & vbCrLf
	Dwt.out "<option value=''>��������ת����</option>" & vbCrLf
	sqlgh="SELECT distinct sb_sscj from sbqt where sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		cjid=cint(rsgh("sb_sscj"))


		sql="SELECT count(sb_id) FROM sbqt WHERE sb_sscj="&cjid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			'i=i+1
			Dwt.out"<option value='"&cjid&"'"
			if cint(request("sscj"))=cjid then Dwt.out" selected"
			sql="SELECT levelname FROM levelname WHERE levelid="&cjid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    cj_name="δ֪��"
			else
			    cj_name=rs("levelname")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&cj_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf

	
	
Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tossgh(this.options[this.selectedIndex].value);}"">" & vbCrLf
Dwt.out "	       <option value=''>��װ����ת����</option>" & vbCrLf



	sqlgh="SELECT distinct sb_ssgh,sb_sscj from sbqt where sb_isdel=false and sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		ghid=cint(rsgh("sb_ssgh"))


		sql="SELECT count(sb_id) FROM sbqt WHERE sb_isdel=false and  sb_ssgh="&ghid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			i=i+1
			Dwt.out"<option value='"&ghid&"'"
			if cint(request("ssgh"))=ghid and request("ssgh")<>"" then Dwt.out" selected"
			
			sql="SELECT gh_name FROM ghname WHERE ghid="&ghid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    gh_name="δ֪��"
			else
			    gh_name=rs("gh_name")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&i&"&nbsp;&nbsp;"&gh_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf


	Dwt.out "</form></Div></Div>" & vbCrLf
	
	
end sub

'********************************************8
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
'URL�д�����
'*******************************************
sub sbshowpage(page,url,total,record)
   Dwt.Out "<Div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      Dwt.Out "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      Dwt.Out ""
   end if 
   if RowCount = 0 and page <>Total then 
     Dwt.Out "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     Dwt.Out ""
   end if
   Dwt.Out"&nbsp;&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&total&"</strong>ҳ&nbsp;&nbsp;"
   if Total =1 then 
    Dwt.Out"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    Dwt.Out"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 Dwt.Out"  <option value='"&page&"' selected >��"&page&"ҳ</option>"
     else
    	 Dwt.Out"  <option value='"&ii&"'>��"&ii&"ҳ</option>"
     end if 
   next 
   
   Dwt.Out" </select>&nbsp;&nbsp;��"&record&"������"
   Dwt.Out "</Div>"
end sub

Sub edit(id,sscj)
    Dwt.Out " <a href=zjtz_qtjc.asp?action=history&id="&id&">ʷ</a>&nbsp;"
    Dwt.Out "<a href=zjtz_qtjc.asp?action=addzjinfo&id="&id&">�ܼ�</a>&nbsp;"
	
if  session("levelclass")=10 then 

	Dwt.Out "<a href=zjtz_qtjc.asp?action=del&id="&id&" onClick=""return confirm('�˲�����ɾ���ñ����е��ܼ��¼��ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
 else
    Dwt.Out "&nbsp;"
 end if 
end Sub

Sub edit2(id,sscj)
    Dwt.Out " <a href=zjtz_qtjc.asp?action=history&id="&id&">ʷ</a>&nbsp;"
if session("levelclass")=sscj then 
    Dwt.Out "<a href=zjtz_qtjc.asp?action=addzjinfo&numb=0&id="&id&">�ܼ�</a>&nbsp;"
    Dwt.Out "<a href=zjtz_qtjc.asp?action=addzjinfo&numb=1&id="&id&">����</a>&nbsp;"
else if session("levelclass")=10 then
    Dwt.Out "<a href=zjtz_qtjc.asp?action=addzjinfo&numb=0&id="&id&">�ܼ�</a>&nbsp;"
    Dwt.Out "<a href=zjtz_qtjc.asp?action=addzjinfo&numb=1&id="&id&">����</a>&nbsp;"
	Dwt.Out "<a href=zjtz_qtjc.asp?action=del&id="&id&" onClick=""return confirm('�˲�����ɾ���ñ����е��ܼ��¼��ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
	else 
    Dwt.Out "&nbsp;"
	end if
end if 
end Sub

Sub delinfo()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from zjinfo_qtjc where id="&id
  rsdel.open sqldel,connzj,1,3
  set rsdel=nothing  
  
  
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end Sub




%>