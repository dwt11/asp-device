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
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->



<%
'���ݿ��� txdate�ֶ�Ϊ�û���ѡֵ��ʱ�䣬txdate1Ϊʵ����д��ʱ�䣬Ĭ������
dim sqlzblog,rszblog,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
classid=request("classid")
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>��Ϣ����ϵͳ--ֵ����־</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/tab.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	

sub add()
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>���ֵ����־</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='zblog.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

    if session("level")=3 then 
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&ssbzh(session("levelzclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='ssbz' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&session("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ��ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out "<select name='isby'><option value='true'>�װ�</option><option value='false'>ҹ��</option></select>"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	 dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=body&style=s_blue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='600' HEIGHT='350'>"
    dwt.out "</iframe>  <input type='hidden' name='body' value=''>"
	'Dwt.OUT "<input type='hidden' name='body' id='body'>"& vbCrLf
    'Dwt.out "<iframe src='neweditor/editor.htm?id=body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"

	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	if session("levelzclass")=17 then 
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 75px'>�����ձ���:</LABEL>"& vbCrLf
		dwt.out"				<DIV class='x-form-element'>"& vbCrLf
		%>
		<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
		  <col  />
		  <col  span="5" />
		  <tr >
			<td  >&nbsp;</td>
		<%dim wlname(9)
		sql="SELECT * from bb_wl where id<>3 and id<>4 and ssbz="&session("levelzclass")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzb,1,1
		if rs.eof and rs.bof then 
			dwt.out  message ("<p align='center'>δ����������������</p>" )
		else
			do while not rs.eof 
			    i=i+1
				wlname(i)=rs("id")
				dwt.out "<td><div align='center'>"&rs("name")&"("&rs("dw")&")</div></td>"	& vbCrLf	
			rs.movenext
		loop
		end if
		dwt.out "<td><div align='center'>��ע</div></td>" %>
		  </tr>
		  <tr >
			<td ><div align="center">0:00-8:00</div></td>
			<%for n=1 to i
			      dwt.out "<td><input name='wl_1_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_1_bz' type='text' size='9'/></td>"
			%>
			
		  </tr>
		  <tr >
			<td ><div align="center">8:00-16:00</div></td>
            <%for n=1 to i
			      dwt.out "<td><input name='wl_2_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_2_bz' type='text' size='9'/></td>"
			%>
			
		  </tr>
		  <tr >
			<td ><div align="center">16:00-24:00</div></td>
            <%for n=1 to i
			      dwt.out "<td><input name='wl_3_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_2_bz' type='text' size='9'/></td>"
			%>
		  </tr>
		</table>
		<%dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 	  
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='wlnumb' type='hidden' value='"&i&"'><input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""location.href='zblog.asp';"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	
	
end sub	

sub saveadd()    
	 
	  '����
   if request("body")<>"" then 
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zblog" 
      rsadd.open sqladd,connzb,1,3
      rsadd.addnew
      rsadd("body")=Trim(request("body"))
      'rsadd("txyear")=year(now())
      'rsadd("txmonth")=month(now())
      'rsadd("txday")=day(now())
      rsadd("userid")=session("userid")
      rsadd("sscj")=session("levelclass")
      rsadd("ssbz")=session("levelzclass")
      rsadd("isby")=request("isby")
      rsadd("txdate")=request("txdate")
		if session("level")<>3 then rsadd("isfc")=true
      if request("isby") then 
	    by="�װ�"
	  else
	    by="ҹ��"
	  end if 	
	  dwt.savesl "ֵ����־","���",request("txdate")&by
	  rsadd.update
      rsadd.close
      set rsadd=nothing
  end if 
	 
	 
	  if session("levelzclass")=17 then 
		  'wlnumb=request("wlnumb")
'		  if request("wl_1_1")<>0 or request("wl_1_2")<>0 or request("wl_1_3")<>0 or request("wl_1_4")<>0  then 
		  'if request("wl_1_1")<>0 or request("wl_1_2")<>0 or request("wl_1_5")<>0   then 

			  set rs2=server.createobject("adodb.recordset")
			  sql2="select * from bb" 
			  rs2.open sql2,connzb,1,3
			  rs2.addnew
				  'rs2("zblog_id")=rsadd("id")
				  rs2("ssbz")=session("levelzclass")
'				  rs2("bansj")=request("wl_1_1")&"/"&request("wl_1_2")&"/"&request("wl_1_3")&"/"&request("wl_1_4")
				  rs2("bansj")=request("wl_1_1")&"/"&request("wl_1_2")&"/"&request("wl_1_5")&"/"&request("wl_1_6")&"/"&request("wl_1_7")&"/"&request("wl_1_8")
				  rs2("banb")=1
				  rs2("bbdate")=request("txdate")
			      RS2("userid")=session("userid")
			  rs2.update
			  rs2.close
'		  end if 

'		  if request("wl_2_1")<>0  or request("wl_2_2")<>0 or request("wl_2_3")<>0 or request("wl_2_4")<>0 then 
'		  if request("wl_2_1")<>0  or request("wl_2_2")<>0 or request("wl_2_5")<>0  then 
			  set rs2=server.createobject("adodb.recordset")
			  sql2="select * from bb" 
			  rs2.open sql2,connzb,1,3
			  rs2.addnew
				  'rs2("zblog_id")=rsadd("id")
				  rs2("ssbz")=session("levelzclass")
'				  rs2("bansj")=request("wl_2_1")&"/"&request("wl_2_2")&"/"&request("wl_2_3")&"/"&request("wl_2_4")
				  rs2("bansj")=request("wl_2_1")&"/"&request("wl_2_2")&"/"&request("wl_2_5")&"/"&request("wl_2_6")&"/"&request("wl_2_7")&"/"&request("wl_2_8")
				  rs2("banb")=2
				  rs2("bbdate")=request("txdate")
			      RS2("userid")=session("userid")
			  rs2.update
			  rs2.close
'		  end if 

'		  if request("wl_3_1")<>0  or request("wl_3_2")<>0 or request("wl_3_3")<>0 or request("wl_3_4")<>0 then 
'		  if request("wl_3_1")<>0  or request("wl_3_2")<>0 or request("wl_3_5")<>0  then 
			  set rs2=server.createobject("adodb.recordset")
			  sql2="select * from bb" 
			  rs2.open sql2,connzb,1,3
			  rs2.addnew
				  'rs2("zblog_id")=rsadd("id")
				  rs2("ssbz")=session("levelzclass")
'				  rs2("bansj")=request("wl_3_1")&"/"&request("wl_3_2")&"/"&request("wl_3_3")&"/"&request("wl_3_4")
				  rs2("bansj")=request("wl_3_1")&"/"&request("wl_3_2")&"/"&request("wl_3_5")&"/"&request("wl_3_6")&"/"&request("wl_3_7")&"/"&request("wl_3_8")
				  rs2("banb")=3
				  rs2("bbdate")=request("txdate")
			      RS2("userid")=session("userid")
			  rs2.update
			  rs2.close
'		  end if 
	  dwt.savesl "������������","���",request("txdate")&by
		  set rs=nothing
	  end if 
		
		
		
		
		
		
		

	  dwt.out "<Script Language=Javascript>location.href='zblog.asp?year="&year(request("txdate"))&"&month="&month(request("txdate"))&"&day="&day(request("txdate"))&"';</Script>"
end sub

sub edit()
     '�༭
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from zblog where id="&id
   rsedit.open sqledit,connzb,1,1
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�༭ֵ����־</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='zblog.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&ssbzh(rsedit("ssbz"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&rsedit("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ��ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("txdate")&"'>"
	dwt.out "<select name='isby'><option value='true' "
	if rsedit("isby")=true then dwt.out "selected"
	dwt.out ">�װ�</option>"
	dwt.out "<option value='false'"
	if rsedit("isby")=false then dwt.out "selected"
	dwt.out ">ҹ��</option></select>"	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ֵ������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	scontent=rsedit("body")
	'Dwt.OUT "<input type='hidden' name='body' id='body' value='"&scontent&"'>"& vbCrLf
    'Dwt.out "<iframe src='neweditor/editor.htm?id=body&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"

	 dwt.out "<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='600' HEIGHT='350'>"
     dwt.out "</iframe><textarea name='body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	if session("levelzclass")=17 then 
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 75px'>�����ձ���:</LABEL>"& vbCrLf
		dwt.out"				<DIV class='x-form-element'>"& vbCrLf
		%>
		<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
		  <col  />
		  <col  span="5" />
		  <tr >
			<td  >&nbsp;</td>
		<%dim wlname(9)
		sql="SELECT * from bb_wl where id<>3 and id<>4 and ssbz="&session("levelzclass")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzb,1,1
		if rs.eof and rs.bof then 
			dwt.out  message ("<p align='center'>δ����������������</p>" )
		else
			do while not rs.eof 
			    i=i+1
				wlname(i)=rs("id")
				dwt.out "<td><div align='center'>"&rs("name")&"("&rs("dw")&")</div></td>"	& vbCrLf	
			rs.movenext
		loop
		end if
		dwt.out "<td><div align='center'>��ע</div></td>" %>
		  </tr>
		  <tr >
			<td ><div align="center">0:00-8:00</div></td>
			<%for n=1 to i
			      dwt.out "<td><input name='wl_1_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_1_bz' type='text' size='9'/></td>"
			%>
			
		  </tr>
		  <tr >
			<td ><div align="center">8:00-16:00</div></td>
            <%for n=1 to i
			      dwt.out "<td><input name='wl_2_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_2_bz' type='text' size='9'/></td>"
			%>
			
		  </tr>
		  <tr >
			<td ><div align="center">16:00-24:00</div></td>
            <%for n=1 to i
			      dwt.out "<td><input name='wl_3_"&wlname(n)&"' type='text' size='9' value='0' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" /></td>"& vbCrLf
			  next 	  
			dwt.out "<td><input name='wl_2_bz' type='text' size='9'/></td>"
			%>
		  </tr>
		</table>
		<%dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 	  








	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'> <input name='id' type='hidden' value='"&request("id")&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""location.href='zblog.asp';"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</DIV>"& vbCrLf
	
	


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from zblog where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connzb,1,3
rsedit("body")=Trim(request("body"))
'rsedit("userid")=session("userid")
rsedit("txdate")=request("txdate")
      rsedit("isby")=request("isby")
rsedit.update
rsedit.close
      if request("isby") then 
	    by="�װ�"
	  else
	    by="ҹ��"
	  end if 	
	  	dwt.savesl "ֵ����־","�༭",request("txdate")&by
	  dwt.out "<Script Language=Javascript>location.href='zblog.asp?year="&year(request("txdate"))&"&month="&month(request("txdate"))&"&day="&day(request("txdate"))&"';</Script>"
	
end sub


sub main()
	url=geturl
	getyear=request("year")
	getmonth=request("month")
	getday=request("day")

    getnowday=date()-1
	
	if getyear="" then getyear=year(getnowday)
	if getmonth="" then getmonth=month(getnowday)
	if getday="" then getday=day(getnowday)



	selectdate=getyear&"-"&getmonth&"-"&getday
	selectdate=cdate(selectdate)
	'message selectdate
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><b>"&selectdate&" ֵ����־</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
	dwt.out "		 <form method='post'  action='zblog.asp'  name='form' >"
	'if session("level")=3 then 
    	dwt.out "		 <a href='/zblog.asp?action=add'>�����־</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	'end if 
	dwt.out "<a href='/zblog.asp?year="&year(selectdate-2)&"&month="&month(selectdate-2)&"&day="&day(selectdate-2)&"'>"&year(selectdate-2)&"��"&month(selectdate-2)&"��"&day(selectdate-2)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<a href='/zblog.asp?year="&year(selectdate-1)&"&month="&month(selectdate-1)&"&day="&day(selectdate-1)&"'>"&year(selectdate-1)&"��"&month(selectdate-1)&"��"&day(selectdate-1)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	dwt.out "<input  type='hidden' name='getyear' value='"&getyear&"' ><input  type='hidden' name='getmonth' value='"&getmonth&"' ><input  type='hidden' name='getday' value='"&getday&"' >		 <select name='year'></select>��<select name='month'></select>��<select name='day'></select>�� &nbsp;&nbsp;<input  type='submit' name='Submit' value=' �鿴 ' style='cursor:hand;'>"
	dwt.out "		 <script type='text/javascript' src='js/selectdate.js'></script>"
	if now()-selectdate>1 then 	dwt.out "<a href='/zblog.asp?year="&year(selectdate+1)&"&month="&month(selectdate+1)&"&day="&day(selectdate+1)&"'>"&year(selectdate-1)&"��"&month(selectdate+1)&"��"&day(selectdate+1)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	if now()-selectdate>2 then 	dwt.out "<a href='/zblog.asp?year="&year(selectdate+2)&"&month="&month(selectdate+2)&"&day="&day(selectdate+2)&"'>"&year(selectdate+2)&"��"&month(selectdate+2)&"��"&day(selectdate+2)&"��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	
	
	dwt.out "	</form></div>"
	dwt.out "</div>"

   
   
   
	

	dwt.out "<div class='navg'>"
	dwt.out "  <div id='system' class='mainNavg'>"
	dwt.out "    <ul>"
	if request("isfc")<>1 then 
		sscjid=request("sscj")
		'101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		if sscjid="" and session("levelclass")<5 then
		 sscjid=session("levelclass")    '101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		else
		   if sscjid="" then sscjid=1    '101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		 end if  
	end if 
	sqlsscj="SELECT * from levelname where levelclass=1 and levelid<5"
	set rssscj=server.createobject("adodb.recordset")
	rssscj.open sqlsscj,conn,1,1
	if rssscj.eof and rssscj.bof then 
		dwt.out  message ("<p align='center'>δ�����������</p>" )
	else
	do while not rssscj.eof 
		if cint(sscjid)=rssscj("levelid") then 
		   dwt.out "<li id='systemNavg'><a href='#'>"&rssscj("levelname")&"</a></li>"
		else
		   dwt.out "<li><a href='zblog.asp?sscj="&rssscj("levelid")&"&year="&getyear&"&month="&getmonth&"&day="&getday&"'>"&rssscj("levelname")&"</a></li>"
		end if    
	rssscj.movenext
	loop
	end if 
    if request("isfc")=1 then 
      dwt.out "<li id='systemNavg'><a href='#'>�ֳ�ֵ��</a></li>"
	else
      dwt.out "<li><a href='zblog.asp?isfc=1&year="&getyear&"&month="&getmonth&"&day="&getday&"'>�ֳ�ֵ��</a></li>"	end if 
	rssscj.close
	set rssscj=nothing
	  
	  
    dwt.out "</ul>"
    dwt.out " </div>"
	
	dwt.out "  <div class='textbody'>"
	if request("isfc")=1 then 
	    call fczb(getyear,getmonth,getday)
    else
		sqlssbz="SELECT * from bzname where sscj="&sscjid
		set rsssbz=server.createobject("adodb.recordset")
		rsssbz.open sqlssbz,conn,1,1
		if rsssbz.eof and rsssbz.bof then 
			dwt.out  message ("<p align='center'>��Ӱ����ſ��������־</p>" )
		else
		
		
		
		
		do while not rsssbz.eof 

				sqlzblog="SELECT * from zblog where sscj="&sscjid&" and ssbz="&rsssbz("id")&" and year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" and isby=true and isfc=false ORDER BY txdate1 aSC "
				set rszblog=server.createobject("adodb.recordset")
				rszblog.open sqlzblog,connzblog,1,1
				if rszblog.eof and rszblog.bof then 
                    dwt.out "<span style='font-size:14px;color:#0000ff;font-weight: bold;'>"&rsssbz("bzname")&"</span><br><br>"
					dwt.out  "<div class='textbody1'>δ��д"&selectdate&"<b>�װ�</b>ֵ����־</div>"
				else
					dwt.out "<span style='font-size:14px;color:#0000ff;font-weight: bold;'>"&rsssbz("bzname")&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"
					dwt.out "<b>"&selectdate&" �װ�</b> ֵ����:<b>"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)&"</b> ����ʱ��:"&rszblog("txdate1")
				do while not rszblog.eof 
	'λ��ʶ���ܣ���ʾ��Ӧλ�ŵļ졢�����ܣ�				
'dwt.out "<div class='textbody1'>"&whsb(DecodeFilter(rszblog("body"),"FONT,STRONG"),sscjid,selectdate)&"&nbsp;&nbsp;&nbsp;&nbsp;"
dwt.out "<div class='textbody1'>"&DecodeFilter(rszblog("body"),"FONT,STRONG")&"&nbsp;&nbsp;&nbsp;&nbsp;"
					call editdel(rszblog("id"),rszblog("sscj"),rszblog("userid"),"zblog.asp?action=edit&id=","zblog.asp?action=del&id=")
				    dwt.out "</div>"
				rszblog.movenext
				loop
				end if 
				rszblog.close	
				
				'ҹ��
				dwt.out "<br/><br/>"
				sqlzblog="SELECT * from zblog where sscj="&sscjid&" and ssbz="&rsssbz("id")&" and year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" and isby=false and isfc=false"
				set rszblog=server.createobject("adodb.recordset")
				rszblog.open sqlzblog,connzblog,1,1
				if rszblog.eof and rszblog.bof then 
					dwt.out  "<div class='textbody1'>δ��д <b>ҹ��</b> ֵ����־</div>"
				else
				dwt.out selectdate&" <b>ҹ��</b> ֵ����:<b>"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)&"</b> ����ʱ��:"&rszblog("txdate1")
				xh=0
				do while not rszblog.eof 
					'dwt.out formatdatetime(rszblog("txdate"),vblongtime)&"&nbsp;&nbsp;&nbsp;&nbsp;"
					'dwt.out "&nbsp;&nbsp;&nbsp;&nbsp;ֵ���ˣ�"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)
					'xh=xh+1
					'dwt.out "<div class='textbody1'>"&whsb(DecodeFilter(rszblog("body"),"FONT,STRONG"),sscjid,selectdate)&"&nbsp;&nbsp;&nbsp;&nbsp;"
dwt.out "<div class='textbody1'>"&DecodeFilter(rszblog("body"),"FONT,STRONG")&"&nbsp;&nbsp;&nbsp;&nbsp;"
					call editdel(rszblog("id"),rszblog("sscj"),rszblog("userid"),"zblog.asp?action=edit&id=","zblog.asp?action=del&id=")
				    dwt.out "</div>"
				rszblog.movenext
				loop
				end if 
				rszblog.close	
				dwt.out "<br/><br/><br/>"
					if rsssbz("id")=17 then
                    if selectdate < #2009-6-18# then%>

								<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
								  <col  />
								  <col  span="5" />
								  <tr >
									<td  >һ��</td>
									<td><div align='center'>����(��)</div></td>
									<td><div align='center'>��(������)</div></td>
									<td><div align='center'>Һ��Һλ(��)</div></td>
									<td><div align='center'>����Һλ(��)</div></td>
									<td>��ע</td>
								  </tr>
								  <tr >
								<td ><div align="center">0:00-8:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=1 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">8:00-16:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=2 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(3)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">16:00-24:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=3 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(3)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if 		
							dwt.out "  </tr>"
				            dwt.out "<div align=center><b>"&selectdate&"����</b></div><div align='right'>����:"&tbname&" ��ʱ��:"&tbdate&"</div>"
							dwt.out "</table>"
                         else %>
				
								<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
								  <col  />
								  <col  span="5" />
								  <tr >
									<td  >һ��</td>
									<td><div align='center'>����(��)</div></td>
									<td><div align='center'>��(������)</div></td>
									<td><div align='center'>������(��)</div></td>
									<td>��ע</td>
								  </tr>
								  <tr >
								<td ><div align="center">0:00-8:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=1 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">8:00-16:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=2 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">16:00-24:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=3 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(0)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(1)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(2)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if 		
							dwt.out "  </tr>"
				            dwt.out "<div align=center><b>"&selectdate&"����</b></div><div align='right'>����:"&tbname&" ��ʱ��:"&tbdate&"</div>"
							dwt.out "</table>"

%>




<br>
<table border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
								  <col  />
								  <col  span="5" />
								  <tr >
									<td  >����</td>
									<td><div align='center'>����(��)</div></td>
									<td><div align='center'>��(������)</div></td>
									<td><div align='center'>������(��)</div></td>
									<td>��ע</td>
								  </tr>
								  <tr >
								<td ><div align="center">0:00-8:00</div></td>
							<%

 on error resume next   '121126ǰû�ж��ڵ����ݣ���������δ���
sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=1 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
                                                          

									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(3)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(4)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(5)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">8:00-16:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=2 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(3)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(4)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(5)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if %>		
							  </tr>
							  <tr >
								<td ><div align="center">16:00-24:00</div></td>
							<%sql="SELECT * from bb where year(bbdate)="&year(selectdate)&" and month(bbdate)="&month(selectdate)&" and day(bbdate)="&day(selectdate)&" and banb=3 order by update desc"
							set rs=server.createobject("adodb.recordset")
							rs.open sql,connzb,1,1
							if rs.eof and rs.bof then 
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>0</div></td>"
									dwt.out "<td><div align='center'>&nbsp;</div></td>"
							else
									bansj=split(rs("bansj"),"/")
									dwt.out "<td><div align='center'>"&bansj(3)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(4)&"</div></td>"
									dwt.out "<td><div align='center'>"&bansj(5)&"</div></td>"
									dwt.out "<td><div align='center'>"&rs("bz")&"&nbsp;</div></td>"
							        tbname=usernameh(rs("userid"))
									tbdate=rs("update")
							end if 		
							dwt.out "  </tr>"
				            
							dwt.out "</table>"
				end if
				end if 
               ' dwt.out "</FIELDSET>"
		rsssbz.movenext
		loop
		end if 
		rsssbz.close	
	end if 
	dwt.out "</div>"
	dwt.out "</div>	"
end sub	

	
	
	
	
	
	
	
	
function fczb(getyear,getmonth,getday)				
	'�װ�
	sqlzblog="SELECT * from zblog where year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" and isby=true and isfc=true ORDER BY txdate1 aSC "
	set rszblog=server.createobject("adodb.recordset")
	rszblog.open sqlzblog,connzblog,1,1
	if rszblog.eof and rszblog.bof then 
		dwt.out  "δ��д"&selectdate&"<b>�װ�</b>ֵ����־"
	else
	dwt.out selectdate&" <b>�װ�</b> ֵ����:"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)&" ����ʱ��:"&rszblog("txdate1")
	dwt.out "<div style='padding-left:20px;padding-top:10px'>"
	do while not rszblog.eof 
		'dwt.out formatdatetime(rszblog("txdate"),vblongtime)&"&nbsp;&nbsp;&nbsp;&nbsp;"
		'dwt.out "&nbsp;&nbsp;&nbsp;&nbsp;ֵ���ˣ�"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)
	
		dwt.out "<div style='padding-left:10px;padding-top:10px;padding-bottom:20px;'>"&DecodeFilter(rszblog("body"),"FONT,STRONG")&"&nbsp;&nbsp;&nbsp;&nbsp;"
		call editdel(rszblog("id"),rszblog("sscj"),rszblog("userid"),"zblog.asp?action=edit&id=","zblog.asp?action=del&id=")
		dwt.out "</div>"
	rszblog.movenext
	loop
	dwt.out "</div>"
	end if 
	rszblog.close	
	
	'ҹ��
	dwt.out "<br/><br/>"
	sqlzblog="SELECT * from zblog where  year(txdate)="&getyear&" and month(txdate)="&getmonth&" and day(txdate)="&getday&" and isfc=true and isby=false"
	set rszblog=server.createobject("adodb.recordset")
	rszblog.open sqlzblog,connzblog,1,1
	if rszblog.eof and rszblog.bof then 
		dwt.out  "δ��д"&selectdate&"<b>ҹ��</b>ֵ����־"
	else
	dwt.out selectdate&" <b>ҹ��</b> ֵ����:"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)&" ����ʱ��:"&rszblog("txdate1")
	dwt.out "<div style='padding-left:20px;padding-top:10px'>"
	xh=0
	do while not rszblog.eof 
		'dwt.out formatdatetime(rszblog("txdate"),vblongtime)&"&nbsp;&nbsp;&nbsp;&nbsp;"
		'dwt.out "&nbsp;&nbsp;&nbsp;&nbsp;ֵ���ˣ�"&conn.Execute("SELECT username1 FROM userid WHERE id="&rszblog("userid"))(0)
		'xh=xh+1
		dwt.out "<div style='padding-left:10px;padding-top:10px;padding-bottom:20px;'>"&DecodeFilter(rszblog("body"),"FONT,STRONG")&"&nbsp;&nbsp;&nbsp;&nbsp;"
		call editdel(rszblog("id"),rszblog("sscj"),rszblog("userid"),"zblog.asp?action=edit&id=","zblog.asp?action=del&id=")
		dwt.out "</div>"
	rszblog.movenext
	loop
	end if 
	rszblog.close	
end function



sub del()
ID=request("ID")
	sqledit="select * from zblog where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connzb,1,1
	if rsedit("isby") then 
	  by="�װ�"
	else
	  by="ҹ��"
	end if   
	    
	dwt.savesl "ֵ����־","ɾ��",rsedit("txdate")&by
	rsedit.close



set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from zblog where id="&id
rsdel.open sqldel,connzb,1,3
dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub

sub editdel(id,sscj,userid,editurl,delurl)
 if session("userid")=userid or session("level")=0 or session("level")=1 and session("levelclass")=sscj then 
    response.write "<a href="&editurl&id&">�༭</a>&nbsp;"
	if session("level")=1 or session("level")=0 and session("levelclass")=sscj then  response.write "<a href="&delurl&id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a>"
 
 end if 
end sub






dwt.out  "</body></html>"

Call CloseConn
%>