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
dwt.out  "<title>��Ϣ����ϵͳ--Ѳ��̨��</title>"& vbCrLf
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
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.pqname.value==''){" & vbCrLf
Dwt.out "      alert('����Ƭ������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.pqname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.whname.value==0){" & vbCrLf
Dwt.out "      alert('Ѳ�������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.whname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.bjrname.value==''){" & vbCrLf
Dwt.out "      alert('������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.bjrname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf

  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>��Ӱ���̨��</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='xjtz.asp' name='form'  onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
   dwt.out"<input name='sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   dwt.out"<select name='ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  dwt.out"<option value='0'>����</option>"
   else   
      do while not rs.eof
	     dwt.out"<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	  end if 
	 dwt.out"</select>" 
  rs.close
  set rs=nothing	
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 95px'>����Ƭ������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=pqname  >"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""pqname"", function() {return ""/inc/autocomplete.asp?dbname=scgldb&zdtext=pqname&btext=xjtz&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	



	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ѳ�������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=whname  >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 95px'>������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=bjrname  >"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""bjrname"", function() {return ""/inc/autocomplete.asp?dbname=scgldb&zdtext=bjrname&btext=xjtz&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf



	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out "<input name='update'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'/>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ѳ�����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out "<select name=orderby><option value=1>1</option><option value=2>2</option><option value=3>3</option><option value=4>4</option><option value=5>5</option></select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			 <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""location.href='zblog.asp';"" style='cursor:hand;'>"& vbCrLf
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
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xjtz" 
      rsadd.open sqladd,connscgl,1,3
      rsadd.addnew
      rsadd("sscj")=request("sscj")
      rsadd("ssbz")=request("ssbz")
      rsadd("pqname")=Trim(request("pqname"))
      rsadd("bjrname")=Trim(request("bjrname"))
      rsadd("whname")=Trim(request("whname"))
      rsadd("update")=request("update")
      rsadd("orderby")=request("orderby")
		
	  dwt.savesl "����̨��","���",Trim(request("whname"))&Trim(request("bjrname"))
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	 
	 
	
		
		
		
		
		
		
		

	  dwt.out "<Script Language=Javascript>location.href='xjtz.asp';</Script>"
end sub

sub edit()
     '�༭
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.pqname.value==''){" & vbCrLf
Dwt.out "      alert('����Ƭ������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.pqname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.whname.value==0){" & vbCrLf
Dwt.out "      alert('Ѳ�������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.whname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.bjrname.value==''){" & vbCrLf
Dwt.out "      alert('������Ϊ�գ�');" & vbCrLf
Dwt.out "   document.form.bjrname.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf

	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from xjtz where id="&id
   rsedit.open sqledit,connscgl,1,1
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�༭Ѳ��̨��</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='xjtz.asp' name='form'  onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
   dwt.out"<input name='sscj' type='text' value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
   dwt.out"<input name='ssbz' type='text' value='"&ssbzh(rsedit("ssbz"))&"'  disabled='disabled' >"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 95px'>����Ƭ������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=pqname  value='"&rsedit("pqname")&"'  disabled='disabled' >"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""pqname"", function() {return ""/inc/autocomplete.asp?dbname=scgldb&zdtext=pqname&btext=xjtz&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	



	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ѳ�������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=whname   value='"&rsedit("whname")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 95px'>������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=bjrname  value='"&rsedit("bjrname")&"'  >"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""bjrname"", function() {return ""/inc/autocomplete.asp?dbname=scgldb&zdtext=bjrname&btext=xjtz&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"

	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf



	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out "<input name='update'  onClick='new Calendar(0).show(this)' readOnly  value='"& rsedit("update") &"'/>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ѳ�����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out "<select name=orderby  disabled='disabled'>"
	dwt.out "<option value=1"
	if  rsedit("orderby")=1 then dwt.out  " selected "
	dwt.out ">1</option>"
	dwt.out "<option value=2"
	if  rsedit("orderby")=2 then dwt.out  " selected "
	dwt.out ">2</option>"
	dwt.out "<option value=3"
	if  rsedit("orderby")=3 then dwt.out  " selected "
	dwt.out ">3</option>"
	dwt.out "<option value=4"
	if  rsedit("orderby")=4 then dwt.out  " selected "
	dwt.out ">4</option>"
	dwt.out "<option value=5"
	if  rsedit("orderby")=5 then dwt.out  " selected "
	dwt.out ">5</option>"
	
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  







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
sqledit="select * from xjtz where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connscgl,1,3
      rsedit("bjrname")=Trim(request("bjrname"))
      rsedit("whname")=Trim(request("whname"))
      rsedit("update")=request("update")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
		
	  dwt.savesl "����̨��","�༭",Trim(request("whname"))&Trim(request("bjrname"))
	  dwt.out "<Script Language=Javascript>location.href='xjtz.asp';</Script>"
	
end sub


sub main()
	url=geturl

	'message selectdate
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>����̨��</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
	'if session("level")=3 then 
    	dwt.out "		 <a href='/xjtz.asp?action=add'>��Ӱ���̨��</a>&nbsp;&nbsp;&nbsp;&nbsp;"
	'end if 
	
	dwt.out "	</div>"
	dwt.out "</div>"

   
   
   
	

	dwt.out "<div class='navg'>"
	dwt.out "  <div id='system' class='mainNavg'>"
	dwt.out "    <ul>"
	'if request("isfc")<>1 then 
		sscjid=request("sscj")
		'101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		if sscjid="" and session("levelclass")<5 then
		 sscjid=session("levelclass")    '101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		else
		   if sscjid="" then sscjid=1    '101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		 end if  
	'end if 
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
		   dwt.out "<li><a href='xjtz.asp?sscj="&rssscj("levelid")&"'>"&rssscj("levelname")&"</a></li>"
		end if    
	rssscj.movenext
	loop
	end if 
  
	  
    dwt.out "</ul>"
    dwt.out " </div>"
	
	dwt.out "  <div class='textbody' style='text-align:center'>"
	
		sqlssbz="SELECT * from bzname where sscj="&sscjid
		set rsssbz=server.createobject("adodb.recordset")
		rsssbz.open sqlssbz,conn,1,1
		if rsssbz.eof and rsssbz.bof then 
			dwt.out  message ("<p align='center'>��Ӱ����ſ��������־</p>" )
		else
		
		
		
		
		do while not rsssbz.eof 
							dwt.out "<span style='font-size:14px;color:#0000ff;font-weight: bold;'>"&rsssbz("bzname")&"</span>&nbsp;&nbsp;&nbsp;&nbsp;"

				sqlzblog="SELECT distinct pqname from xjtz where sscj="&sscjid&" and ssbz="&rsssbz("id")
				set rszblog=server.createobject("adodb.recordset")
				rszblog.open sqlzblog,connscgl,1,1
				if rszblog.eof and rszblog.bof then 
					dwt.out  "<div class='textbody1'>δ�ҵ���¼</div>"
				else
		%>
				     <table width="80%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" >Ƭ������</td>
    <td align="center">1��Ѳ���</td>
    <td align="center">2��Ѳ���</td>
    <td align="center">3��Ѳ���</td>
    <td align="center">4��Ѳ���</td>
    <td align="center">5��Ѳ���</td>
  </tr>			
				<%do while not rszblog.eof 
					%>
                    
                    
                    
                    
                    
                    
               
  <tr>
    <td><%=rszblog("pqname")%></td>
    <%for orderid=1 to 5
		sqlzblog1="SELECT * from xjtz where sscj="&sscjid&" and ssbz="&rsssbz("id")&" and pqname='"&rszblog("pqname")&"' and orderby="&orderid
				set rszblog1=server.createobject("adodb.recordset")
				rszblog1.open sqlzblog1,connscgl,1,1
				if rszblog1.eof and rszblog1.bof then 
					dwt.out "<td>&nbsp;</td>"
				else
					dwt.out "<td>"&rszblog1("whname")&"&nbsp;&nbsp;"&rszblog1("bjrname")&" "
					 call editdel(rszblog1("id"),rszblog1("sscj"),"xjtz.asp?action=edit&id=","xjtz.asp?action=del&id=")
					 dwt.out "</td>"
				end if 
	next 
	
	
	
	
	%>
    
  
  
  </tr>
 
                    
                    
                    
                    
                    
                    
                    
                    <%
					
					
					
					
				rszblog.movenext
				loop
				end if 
				rszblog.close	
				
			%>	
				
</table>
				
		<br/>		
				<%
               ' dwt.out "</FIELDSET>"
		rsssbz.movenext
		loop
		end if 
		rsssbz.close	
	
	dwt.out "</div>"
	dwt.out "</div>	"
end sub	

	
	
	
	



sub del()
ID=request("ID")



set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from xjtz where id="&id
rsdel.open sqldel,connscgl,1,3
dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	dwt.savesl "Ѳ��̨��","ɾ��",id
'rsdel.close
set rsdel=nothing  

end sub

sub editdel(id,sscj,editurl,delurl)
 if session("level")=0 or session("level")=1 and session("levelclass")=sscj then 
    response.write "<a href="&editurl&id&">��</a>&nbsp;"
	response.write "<a href="&delurl&id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
 
 end if 
end sub






dwt.out  "</body></html>"

Call CloseConn
%>