<%@language=vbscript codepage=936 %>
<%
Option Explicit
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
dim rs,sql,leftnumb
leftmdb="ybdata/left.mdb"
Set connleft = Server.CreateObject("ADODB.Connection")
connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
connleft.Open connl    

dwt.pagetop "ϵͳ��Ŀ����"
select case Request("action")
  case ""
      call mainclass'��ҳ����ʾ������
  case "mainclass"
      call mainclass'��ҳ����ʾ������
  case "main"
      call main'������
  case "addclass"
      call addclass '���Ӹ�����
  case "saveaddclass"
      call saveaddclass    '���游����
  case "editclass"
      call editclass '�༭������
  case "saveeditclass"
      call saveeditclass '�༭���游����
  case "isbiglevel"
      call isbiglevel '�̳и�����Ȩ��
  case "unisbiglevel"
      call unisbiglevel 'ȡ���̳�
  case "delclass"
      call delclass  'ɾ����������Ϣ
  case "isshartcut"
      	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from left_class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connleft,1,3
       'on error resume next
	  rsedit("isshartcut")=request("isshartcut")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
 
	  
end select	  

sub delclass()
  dim id,sqldel,rsdel
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from left_class where id="&id
  rsdel.open sqldel,connleft,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub

sub addclass()'��ӷ���
	dwt.out "<SCRIPT language=javascript>" & vbCrLf
	dwt.out "function checkadd(){" & vbCrLf
	dwt.out "  if(document.form1.name.value==''){" & vbCrLf
	dwt.out "      alert('��Ŀ����δ��д��');" & vbCrLf
	dwt.out "  document.form1.name.focus();" & vbCrLf
	dwt.out "      return false;" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "</SCRIPT>" & vbCrLf
	dwt.out"<DIV style='WIDTH: 360px;padding-top:100px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�����Ŀ</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='config.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				 <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=name>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf



	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�����ַ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				 <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name='url' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='class' size='1' style='WIDTH: 175px'>"& vbCrLf
	dwt.out "<option value=0>ѡ����������</option>"
	sql="SELECT * from left_class where zclass=0 order by orderby aSC"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connleft,1,1
	if rs.eof and rs.bof then 
 	else
		do while not rs.eof 
			dwt.out "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
		rs.movenext
		loop
	end if 
	rs.close
	set rs=nothing
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"							<DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=orderby>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
				  
	dwt.out"			<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>�Ƿ��ݲ˵���ʾ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='isput' >"& vbCrLf
	dwt.out"<option value='true'>��ʾ</option>"	
	dwt.out"<option value='false'>����ʾ</option>"
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>�Ƿ�Ӻ���ʾ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='isshartcut' >"& vbCrLf
	dwt.out"<option value='true'>�Ӻ�</option>"	
	dwt.out"<option value='false'>���Ӻ�</option>"
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveaddclass'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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

sub saveaddclass()    
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from left_class" 
      rsadd.open sqladd,connleft,1,3
      rsadd.addnew
       'on error resume next
      if request("class")=0 then 
	     rsadd("zclass")=0
      else
		 rsadd("zclass")=request("class")
      end if 
	  rsadd("name")=request("name")
	  rsadd("isput")=request("isput")
	  rsadd("isshartcut")=request("isshartcut")
	  dim orderby
	  orderby=request("orderby")
	  if orderby="" then orderby=0
	  rsadd("orderby")=orderby
	  rsadd("url")=request("url")
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='config.asp';</Script>"
end sub
 



sub saveeditclass()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from left_class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connleft,1,3
       'on error resume next
      if request("class")=0 then 
	     rsedit("zclass")=0
      else
		 rsedit("zclass")=request("class")
      end if 
	  rsedit("name")=request("name")
	  rsedit("isput")=request("isput")
	  rsedit("isshartcut")=request("isshartcut")
	  dim orderby
	  orderby=request("orderby")
	  if orderby="" then orderby=0
	  rsedit("orderby")=orderby
	  rsedit("url")=request("url")

		  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub editclass()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from left_class where id="&id
   rsedit.open sqledit,connleft,1,1
	dwt.out "<SCRIPT language=javascript>" & vbCrLf
	dwt.out "function checkadd(){" & vbCrLf
	dwt.out "  if(document.form1.name.value==''){" & vbCrLf
	dwt.out "      alert('��Ŀ����δ��д��');" & vbCrLf
	dwt.out "  document.form1.name.focus();" & vbCrLf
	dwt.out "      return false;" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "</SCRIPT>" & vbCrLf
	dwt.out"<DIV style='WIDTH: 360px;padding-top:100px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�༭��Ŀ</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='config.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				 <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name='name' value='"&rsedit("name")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�����ַ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				 <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name='url' value='"&rsedit("url")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='class' size='1' style='WIDTH: 175px'>"& vbCrLf
	dwt.out "<option value=0>ѡ����������</option>"
	sql="SELECT * from left_class where zclass=0 order by orderby aSC"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connleft,1,1
	if rs.eof and rs.bof then 
 	else
		do while not rs.eof 
			dwt.out "<option value='"&rs("id")&"'"
			if rsedit("zclass")=rs("id") then dwt.out "selected"
			dwt.out ">"&rs("name")&"</option>"
		rs.movenext
		loop
	end if 
	rs.close
	set rs=nothing
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"							<DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=orderby value='"&rsedit("orderby")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
				  
	dwt.out"			<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>�Ƿ��ݲ˵���ʾ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='isput' >"& vbCrLf
	dwt.out"<option value='true'"
    if rsedit("isput")=true then dwt.out "selected"
	dwt.out ">��ʾ</option>"	
	dwt.out"<option value='false'"
    if rsedit("isput")=false then dwt.out "selected"
	dwt.out ">����ʾ</option>"	
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>�Ƿ�Ӻ���ʾ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <select name='isshartcut' >"& vbCrLf
	dwt.out"<option value='true'"
    if rsedit("isshartcut")=true then dwt.out "selected"
	dwt.out ">�Ӻ�</option>"	
	dwt.out"<option value='false'"
    if rsedit("isshartcut")=false then dwt.out "selected"
	dwt.out ">���Ӻ�</option>"	
	dwt.out "</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveeditclass'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' �� ��' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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


sub isbiglevel()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from left_class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connleft,1,3
       'on error resume next
      
	  rsedit("isbiglevel")=true

	rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub
sub unisbiglevel()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from left_class where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connleft,1,3
       'on error resume next
      
	  rsedit("isbiglevel")=false
	rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


Sub mainclass()
	dwt.out "<SCRIPT language=javascript1.2>" & vbCrLf
	dwt.out "function showsubmenu(sid){" & vbCrLf
	dwt.out "      	 var ss='xxx'+sid;" & vbCrLf
	dwt.out "    whichEl = eval('info' + sid);" & vbCrLf
	dwt.out "    if (whichEl.style.display == 'none'){" & vbCrLf
	dwt.out "        eval(""info"" + sid + "".style.display='block';"");" & vbCrLf
	dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i6.gif' />"";" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "    else{" & vbCrLf
	dwt.out "        eval(""info"" + sid + "".style.display='none';"");" & vbCrLf
	dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i7.gif' />"";" & vbCrLf
	dwt.out "    }" & vbCrLf
	dwt.out "}" & vbCrLf
	dwt.out "</SCRIPT>" & vbCrLf
	
	
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>ϵͳ��Ŀ����</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	dwt.out "<div class='x-toolbar'>" & vbCrLf
	dwt.out "<div align=left><a href=""config.asp?action=addclass"">�����Ŀ</a></div>" & vbCrLf
	dwt.out "</div>"

  
  
  
  sql="SELECT * from left_class where zclass=0 order by orderby aSC"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connleft,1,1
  if rs.eof and rs.bof then 
     message "���κ���Ŀ" 
  else
	 dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
     dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     dwt.out "<tr class=""x-grid-header"">"
     dwt.out "<td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��Ŀ����</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��Ŀ��ַ</div></td>"
     'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>������Ŀ</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�Ƿ��ݲ˵���ʾ</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�Ƿ�����ʾ</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�� ��</div></td>"
     dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>ѡ ��</div></td>"
     dwt.out "    </tr>"
  
  do while not rs.eof 
			dim xh,xh_id
			
			xh=xh+1
			
			if xh mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><a href='#' onclick=""showsubmenu("&rs("id")&");"" id=xxx"&rs("id")&"><img src='/img_ext/i7.gif' /></a>"&xh&"</div></td>"
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rs("name")&"&nbsp;</div></td>"
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rs("url")&"&nbsp;</div></td>"
			'dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">һ��</div></td>"
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rs("isput")&"</div></td>"
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rs("isshartcut")&"</div></td>"  '�����Ƿ���RIGHT��ASPҳ������ʾ
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rs("orderby")&"&nbsp;</div></td>"
		   dwt.out "<td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">&nbsp;"
		   if rs("id")<>15 then
		    if rs("isbiglevel") then 
			    dwt.out "<a href=config.asp?action=unisbiglevel&id="&rs("id")&">ȡ���̳�Ȩ��</a>&nbsp;&nbsp;"
			else	
			    dwt.out "<a href=config.asp?action=isbiglevel&id="&rs("id")&">�ӷ���̳и�����Ȩ��</a>&nbsp;&nbsp;"
			end if 
			dwt.out "<a href=config.asp?action=editclass&id="&rs("id")&">�༭</a>&nbsp;&nbsp;"
		    dwt.out "<a href=config.asp?action=delclass&id="&rs("id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
		   end if 
		   dwt.out "</div></td></tr>"
	    			'����
			dwt.out "<tr ><td  colspan=7 style='display:none' id='info"&rs("id")&"'>"	
			dwt.out "<table width=""90%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"	
			dim sqlz,rsz
			sqlz="SELECT * from left_class where zclass="&rs("id")&" order by orderby aSC"& vbCrLf
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,connleft,1,1
			if rsz.eof and rsz.bof then 
			else
				
				     dwt.out "<tr class=""x-grid-header"">"
				 dwt.out "<td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��Ŀ����</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>��Ŀ��ַ</div></td>"
				 'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>������Ŀ</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�Ƿ�������ʾ</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�Ƿ���ҳ������ʾ</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�� ��</div></td>"
				 dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>ѡ ��</div></td>"
				 dwt.out "    </tr>"

				dim xhz
				xhz=0
				do while not rsz.eof
					'xh=xh+1
					
					xhz=xhz+1
					if xh mod 2 =1 then 
					  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&xh&"-"&xhz&"</div></td>"
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsz("name")&"&nbsp;</div></td>"
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsz("url")&"&nbsp;</div></td>"
'					dwt.out " <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("name")&"-����</div></td>"
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsz("isput")&"&nbsp;</div></td>"
					'����ʾҳ��ֱ�������Ƿ�������ʾ080402
					dim trueorfasle
					if rsz("isshartcut") then 
					    trueorfasle=false
					else
					    trueorfasle=true
					end if 
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><a href='config.asp?action=isshartcut&isshartcut="&trueorfasle&"&id="&rsz("id")&"'>"&rsz("isshartcut")&"</a></div></td>"
					dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("orderby")&"-"&rsz("orderby")&"&nbsp;</div></td>"
				   dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;"
				   if rsz("id")<>16 then
				    dwt.out "<a href=config.asp?action=editclass&id="&rsz("id")&">�༭</a>&nbsp;&nbsp;"
				    dwt.out "<a href=config.asp?action=delclass&id="&rsz("id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
				   end if 
				   dwt.out "</div></td></tr>"
				
				rsz.movenext
				loop
			
			end if 	
			   dwt.out "</table></td></tr>"
			rsz.close
			set rsz=nothing
		
    rs.movenext
    loop
     dwt.out "</table></div>"
end if 
  rs.close
  set rs=nothing
  dwt.out "</div>"
end sub

dwt.out "</body></html>"



'ȡ�ֶε�����
function sbtable_body(sbclass_id,sbtable_name)
dim sql,rs
 sql="SELECT sbtable_body from sbtable where sbtable_sbclassid="&sbclass_id&" and sbtable_name='"&sbtable_name&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then 
     sbtable_body= null
  else
     sbtable_body=rs("sbtable_body")
  end if
end function


'ȡ�ֶ�����˳��
function sbtable_orderby(sbclass_id,sbtable_name)
dim sql,rs
 sql="SELECT sbtable_orderby from sbtable where sbtable_sbclassid="&sbclass_id&" and sbtable_name='"&sbtable_name&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs.eof and rs.bof then 
     sbtable_orderby= 0 
  else
     sbtable_orderby=rs("sbtable_orderby")
  end if
end function
connleft.close
set connleft=nothing%>