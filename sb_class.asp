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
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql

'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>�豸�������ҳ</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out "  if(document.form.sbclass_name.value==''){" & vbCrLf
Dwt.out "      alert('��������δ��д��');" & vbCrLf
Dwt.out "  document.form.sbclass_name.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
select case Request("action")
  case ""
      call mainclass'��ҳ����ʾ������
  case "mainclass"
      call mainclass'��ҳ����ʾ������
  case "main"
      call main'������
  case "sbclass_zq"
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbclass where sbclass_id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conn,1,3
      rsedit("sbclass_zq")=Request("is")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"

  case "addclass"
      call addclass '���Ӹ�����
  case "saveaddclass"
      call saveaddclass    '���游����
  case "editclass"
      call editclass '�༭������
  case "saveeditclass"
      call saveeditclass '�༭���游����
  case "delclass"
      call delclass  'ɾ����������Ϣ
  case "edittable"
      call edittable  'ɾ����������Ϣ
  case "saveedittable"
      call saveedittable  'ɾ����������Ϣ
 
 '�༭�������ݺ͹�������
  case "editjx"
      call editjx  
  case "saveeditjx"
      call saveeditjx  
end select	  



sub addclass()'��ӷ���
   Dwt.out"<form method='post' action='sb_class.asp' name='form' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�豸�����������ӷ���</strong></Div></td>    </tr>"
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"      
    Dwt.out"<td width='88%' class='tdbg'>"
       Dwt.out"<input name='sbclass_name' type='text'></td></tr>"& vbCrLf

    dim rs,sql,rsz,sqlz
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ࣺ </strong></td>"      
    Dwt.out"<td width='88%' class='tdbg'>"

Dwt.out "<select name='sb_class' size='1' id='cat1' onChange=""selectpc(this.value,'b',document.form.sb_zclass)"">"
Dwt.out "  <option selected value='0'>ѡ��һ������</option>"
	sql="SELECT * from sbclass where sbclass_zclass=0 "& vbCrLf
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    do while not rs.eof
       	Dwt.out"<option value='"&rs("sbclass_id")&"'>"&rs("sbclass_name")&"</option>"& vbCrLf
		rs.movenext
	loop
	rs.close
	set rs=nothing
	Dwt.out "</select>"
	Dwt.out "<select name='sb_zclass' size='1' id='cat2' >"
	Dwt.out "  <option selected value=0>ѡ���������</option>"
	Dwt.out "</select></td></tr>"& vbCrLf
	Dwt.out "<script language='javascript'>"& vbCrLf
	Dwt.out "function selectpc(parentValue,child,addObj){"& vbCrLf


dim b,bv,b_p
	sql="SELECT * from sbclass where sbclass_zclass=0 "& vbCrLf
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
         b="var b =   new Array("
        bv="var bv =   new Array("
        b_p="var b_p =   new Array("
   
	do while not rs.eof
		sqlz="SELECT * from sbclass where sbclass_zclass="&rs("sbclass_id")
        set rsz=server.createobject("adodb.recordset")
        rsz.open sqlz,conn,1,1
        if rsz.eof and rsz.bof then
		   b=b&"'�޶�������',"
		   bv=bv&"'0',"
		   b_p=b_p&"'"&rs("sbclass_id")&"',"
		else
		do while not rsz.eof
			b=b&"'"&rsz("sbclass_name")&"',"
			bv=bv&"'"&rsz("sbclass_id")&"',"
			b_p=b_p&"'"&rs("sbclass_id")&"',"


		   rsz.movenext
	    loop
	    end if 
		rsz.close
	    set rsz=nothing

		rs.movenext
	loop
	rs.close
	set rs=nothing
	b=left(b,len(b)-1)
	bv=left(bv,len(bv)-1)
	b_p=left(b_p,len(b_p)-1)
	b=b&");"
	bv=bv&");"
	b_p=b_p&");"
	Dwt.out b & vbCrLf
	Dwt.out bv & vbCrLf
	Dwt.out b_p & vbCrLf
	
	
	
	Dwt.out "var labelValue = new Array();"& vbCrLf
	Dwt.out "var labelText =  new Array();"& vbCrLf
	Dwt.out "var k = 0;"& vbCrLf
	
	Dwt.out "cObj = eval(child);"& vbCrLf
	Dwt.out "cObjV = eval(child+'v');"& vbCrLf
	Dwt.out "cpObj = eval(child + '_p');"& vbCrLf
	Dwt.out "for(i=0; i<cpObj.length; i++)"& vbCrLf
	Dwt.out "{"& vbCrLf
	Dwt.out "	if(cpObj[i] == parentValue)"& vbCrLf
	Dwt.out "	{"& vbCrLf
	Dwt.out "		labelText[k] =  cObj[i];"& vbCrLf
	Dwt.out "		labelValue[k] =	cObjV[i]; "& vbCrLf
	Dwt.out "		k++;"& vbCrLf
	Dwt.out "	}"& vbCrLf
	Dwt.out "}"& vbCrLf
	
	
	Dwt.out "addObj.options.length = 0;"& vbCrLf
	Dwt.out "addObj.options[0] = new Option('==ѡ���������==','0');"& vbCrLf
	Dwt.out "for(i = 0; i < labelText.length; i++) {"& vbCrLf
	Dwt.out "	addObj.add(document.createElement('option'));"& vbCrLf
	Dwt.out "	addObj.options[i+1].text=labelText[i];"& vbCrLf
	Dwt.out "	addObj.options[i+1].value=labelValue[i];"& vbCrLf
	Dwt.out "}"& vbCrLf
	Dwt.out "addObj.selectedIndex = 0;"& vbCrLf
Dwt.out "}"& vbCrLf
Dwt.out "</script>"& vbCrLf
	
	
	
	
	 
		 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbclass_orderby' type='text'></td></tr>"& vbCrLf
   
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveaddclass'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
	message "ʲô����ѡ������һ������;ֻѡһ�����࣬��ѡ�������࣬���������һ�������µ��ӷ���<br>���ӷ��������أ����Ӻ�ֻ���޸�����������ʾ�������޸����������ϼ�����"
end sub	

sub saveaddclass()    
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from sbclass" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
       'on error resume next
      if request("sb_class")=0 then 
	     rsadd("sbclass_zclass")=0
      else
	     if request("sb_zclass")=0 then 
		    rsadd("sbclass_zclass")=request("sb_class")
		 else
		    rsadd("sbclass_zclass")=request("sb_zclass")
		 end if 
      end if 
	  rsadd("sbclass_name")=request("sbclass_name")
	  rsadd("sbclass_isput")=true
	  rsadd("sbclass_orderby")=0
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>location.href='sb_class.asp?action=mainclass';</Script>"
end sub
 



'sub saveedit()    
'	  '����
'	  dim rsedit,sqledit
'      set rsedit=server.createobject("adodb.recordset")
'      sqledit="select * from sbclass where sbclass_id="&ReplaceBadChar(Trim(request("ID")))
'      rsedit.open sqledit,conn,1,3
'      rsedit("sbclass_name")=Trim(Request("sbclass_name"))
'      rsedit.update
'      rsedit.close
'      set rsedit=nothing
'	  Dwt.out"<Script Language=Javascript>history.go(-2)<Script>"
'end sub
sub saveeditclass()    
	  '����
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbclass where sbclass_id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conn,1,3
      rsedit("sbclass_name")=Trim(Request("sbclass_name"))
      	  rsedit("sbclass_isput")=request("sbclass_isput")
	  	  rsedit("sbclass_orderby")=request("sbclass_orderby")
		  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub saveedittable()
	  dim rsedit,sqledit,tablei
     
   for tablei=1 to 30
	if Trim(Request("sbtbale_c"&tablei))="" then 
	  dim rsdel,sqldel
	  set rsdel=server.createobject("adodb.recordset")
      sqldel="delete * from sbtable where sbtable_sbclassid="&ReplaceBadChar(Trim(request("ID")))&" and sbtable_name='sb_c"&tablei&"'"
      rsdel.open sqldel,conn,1,3
	  set rsdel=nothing
	else
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbtable where sbtable_sbclassid="&ReplaceBadChar(Trim(request("ID")))&" and sbtable_name='sb_c"&tablei&"'"
      rsedit.open sqledit,conn,1,3
      if rsedit.eof then 
		rsedit.addnew
		rsedit("sbtable_name")="sb_c"&tablei
		rsedit("sbtable_sbclassid")=request("ID")
		rsedit("sbtable_body")=Trim(Request("sbtbale_c"&tablei))
		rsedit("sbtable_orderby")=Request("sbtable_orderby"&tablei)
		rsedit.update
	  else
		  'rsedit("sbtable_name")=Trim(Request("sbtable_name"))
		  rsedit("sbtable_body")=Trim(Request("sbtbale_c"&tablei))
		  rsedit("sbtable_orderby")=Request("sbtable_orderby"&tablei)
		  rsedit.update
	  end if 
      rsedit.close
      set rsedit=nothing
    end if 	
  next 
	
	for tablei=23 to 23
	if Trim(Request("sbtbale_b"&tablei))="" then 
	  'dim rsdel,sqldel
	  set rsdel=server.createobject("adodb.recordset")
      sqldel="delete * from sbtable where sbtable_sbclassid="&ReplaceBadChar(Trim(request("ID")))&" and sbtable_name='sb_b"&tablei&"'"
      rsdel.open sqldel,conn,1,3
	  set rsdel=nothing
	else  
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sbtable where sbtable_sbclassid="&ReplaceBadChar(Trim(request("ID")))&" and sbtable_name='sb_b"&tablei&"'"
      rsedit.open sqledit,conn,1,3
      if rsedit.eof then 
		rsedit.addnew
		rsedit("sbtable_name")="sb_b"&tablei
		rsedit("sbtable_sbclassid")=request("ID")
		rsedit("sbtable_body")=Trim(Request("sbtbale_b"&tablei))
		rsedit("sbtable_orderby")=Request("sbtable_orderby"&tablei)
		rsedit.update
	  else
		  'rsedit("sbtable_name")=Trim(Request("sbtable_name"))
		  rsedit("sbtable_body")=Trim(Request("sbtbale_b"&tablei))
		  rsedit("sbtable_orderby")=Request("sbtable_orderby"&tablei)
		  rsedit.update
	  end if 
      rsedit.close
      set rsedit=nothing
	end if 
	 next 
	 
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"

end sub

sub delclass()
dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from sbclass where sbclass_id="&request("id")
  rsdel.open sqldel,conn,1,3
  Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub edittable()
	dim id,rsedit,sqledit
	id=ReplaceBadChar(Trim(request("id")))
	Dwt.out"<form method='post' action='sb_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.out"<tr class='title'><td height='22' colspan='2'>"
	Dwt.out"<Div align='center'><strong>�༭"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&id)(0)&"�������</strong></Div></td></tr></table>"
	
	Dwt.out"<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr class='title'>"   & vbCrLf   
	 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>���</strong></Div></td>"
     Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>����</strong></Div></td>"
     Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>����(�������ִ��ظ�)</strong></Div></td>"
     Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>����</strong></Div></td>"
     Dwt.out "</tr>"
   dim tablei,sbtable_name
    for tablei=1 to 30
        Dwt.out " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
	    sbtable_name="sb_c"&tablei
		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><Div align='center'>"&tablei&"</Div></td>"
		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><input name='sbtbale_c"&tablei&"' type='text' value='"&sbtable_body(id,sbtable_name)&"'></td>"
		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><input name='sbtable_orderby"&tablei&"' type='text' value='"&sbtable_orderby(id,sbtable_name)&"'></td>"
		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"">�ı�</td>"
	Dwt.out "</tr>"
	next
'	for tablei=23 to 23
'        Dwt.out " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
'	    sbtable_name="sb_b"&tablei
'		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><Div align='center'>"&tablei&"</Div></td>"
'		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><input name='sbtbale_b"&tablei&"' type='text' value='"&sbtable_body(id,sbtable_name)&"'></td>"
'		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><input name='sbtable_orderby"&tablei&"' type='text' value='"&sbtable_orderby(id,sbtable_name)&"'></td>"
'		Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"">��/��</td>"
'	Dwt.out "</tr>"
'	next

	Dwt.out"</table><table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedittable'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
end sub


sub editjx()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from sbclass where sbclass_id="&id
   rsedit.open sqledit,conn,1,1
   
   dim sbclassjxnr,sbclassjxgzxx
   
   sbclassjxnr=rsedit("sb_jxnr_class")
   sbclassjxgzxx=rsedit("sb_jxgzxx_class")
    dwt.out"<form method='post' action='sb_class.asp' name='form1' >"
   	dwt.out"<div align=center><br><br>�༭�豸����<b>"&rsedit("sbclass_name")&"</b>�ļ�������<DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf



	dwt.out" <table width='550' border='1'>"& vbCrLf
	dwt.out" <tr>    <td colspan='2' align=center><b>��������</b></td>  </tr>"& vbCrLf
 
 
 
 
 
  dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxgzxx where sbjxgzxx_zclass=0 order by  sbjxgzxx_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
  
  do while not rsbody.eof 
 
 
 	dwt.out"			 <tr><td align=right><b>"&rsbody("sbjxgzxx_name")&"</b>:</td><td>"& vbCrLf
							'����
					sqlz="SELECT * from sbjxgzxx where sbjxgzxx_zclass="&rsbody("sbjxgzxx_id")&" order by  sbjxgzxx_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
						
						do while not rsz.eof
						
			dwt.out"<input type='checkbox' name='jxgzxx' value='"&rsz("sbjxgzxx_id")&"'" 
				call checkbox(sbclassjxgzxx,rsz("sbjxgzxx_id"))
							Dwt.out ">"	
							dwt.out rsz("sbjxgzxx_name") & "<br>"
							
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing
			dwt.out"</td></tr>"& vbCrLf

		
    rsbody.movenext
    loop
end if 
  rsbody.close
  set rsbody=nothing
	
	
	
	

				  
	
	dwt.out"		</table>"& vbCrLf






	
	

 
 	dwt.out" <BR><BR><table width='550' border='1'>"& vbCrLf
	dwt.out" <tr>    <td colspan='2' align=center><b>��������</b></td>  </tr>"& vbCrLf

 
 
  sqlbody="SELECT * from sbjxnr where sbjxnr_zclass=0 order by  sbjxnr_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
  
  do while not rsbody.eof 
 
  	dwt.out"			 <tr><td align=right><b>"&rsbody("sbjxnr_name")&"</b>:</td><td>"& vbCrLf

							'����
					sqlz="SELECT * from sbjxnr where sbjxnr_zclass="&rsbody("sbjxnr_id")&" order by  sbjxnr_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					
					dwt.out "&nbsp;"
					else
						
						do while not rsz.eof
						
			dwt.out"<input type='checkbox' name='jxnr' value='"&rsz("sbjxnr_id")&"'" 
				call checkbox(sbclassjxnr,rsz("sbjxnr_id"))
							Dwt.out ">"	
							dwt.out rsz("sbjxnr_name") & "<br>"
							
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing
			dwt.out"</td></tr>"& vbCrLf
		
    rsbody.movenext
    loop
end if 
  rsbody.close
  set rsbody=nothing
	
	
	
	

				  
	
	dwt.out"			</TABLE>"& vbCrLf
	
		dwt.out"		<br><br><br><br>	  <input name='action' type='hidden' value='saveeditjx'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf

	
	
	dwt.out"	<br><br><br><br></div> "& vbCrLf  	
	dwt.out"		  </FORM>"& vbCrLf
       rsedit.close
       set rsedit=nothing
end sub

sub saveeditjx()
	dim id,rsedit,sqledit
	id=request("id")
	'For i = LBound(checkuser) To UBound(checkuser)
		set rsedit=server.createobject("adodb.recordset")
		sqledit="select * from sbclass where sbclass_ID="&id
		rsedit.open sqledit,conn,1,3
        'message Request("check_display")&"/"&Request("check_new")&"/"&Request("check_edit")&"/"&Request("check_del")
		rsedit("sb_jxgzxx_class")=Request("jxgzxx")
		rsedit("sb_jxnr_class")=Request("jxnr")
		rsedit.update
		rsedit.close
	'Next 
	
	Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"


end sub


'�������ƣ�checkbox ҳ���Ƿ�ѡ��
'���ã��ж��豸����ļ������ݺ͹��������Ƿ�ѡ�� ���ݿ��������������checked
Function checkbox(sbclassjx,jxid)
	dim sbclassjx1,i
	if not isnull( sbclassjx ) then 
	  sbclassjx1=split(sbclassjx,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=cint(jxid) then dwt.out " checked "
	 Next 
	end if  
end Function







sub editclass()
dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from sbclass where sbclass_id="&id
   rsedit.open sqledit,conn,1,1
   Dwt.out"<form method='post' action='sb_class.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�༭����</strong></Div></td>    </tr>"
     
     Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbclass_name' type='text' value='"&rsedit("sbclass_name")&"'></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�Ƿ������˵���ʾ�� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='70%' class='tdbg'>"
	Dwt.out"<select name='sbclass_isput' size='1' >"
	Dwt.out"<option value='true'"
	if rsedit("sbclass_isput")=true then Dwt.out" selected" 
	Dwt.out">��ʾ</option>"
	Dwt.out"<option value='false' "
	if rsedit("sbclass_isput")=false then Dwt.out"selected"
	Dwt.out">����ʾ</option>"
	Dwt.out"</select></td></tr>"
     Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>���� </strong></td>"   & vbCrLf   
     Dwt.out"<td width='88%' class='tdbg'><input name='sbclass_orderby' type='text' value='"&rsedit("sbclass_orderby")&"'></td></tr>"& vbCrLf

		Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveeditclass'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
	
       rsedit.close
       set rsedit=nothing
end sub

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



Sub mainclass()
  	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�豸�������---�������</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
	Dwt.out "<Div class='x-toolbar'>" & vbCrLf
	Dwt.out "<Div align=left><a href=""sb_class.asp?action=addclass"">��ӷ���</a></Div>" & vbCrLf
	Dwt.out "</Div>"

  dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbclass where sbclass_zclass=0 order by  sbclass_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
  	 Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
     Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
     Dwt.out "<tr class=""x-grid-header"">"
     Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�Ƿ������˵�����ʾ</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�� ��</Div></td>"
     Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>ѡ ��</Div></td>"
     Dwt.out "    </tr>"
  
  do while not rsbody.eof 
	  dim xh,xh_id
		xh=xh+1
			if xh mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
        Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"</Div></td>"
        Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><b>"&rsbody("sbclass_name")&"</b>&nbsp;</Div></td>"
        Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">һ��</Div></td>"
        Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("sbclass_isput")&"&nbsp;</Div></td>"
        Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("sbclass_orderby")&"&nbsp;</Div></td>"
       Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
	   Dwt.out "<a href=sb_class.asp?action=editclass&id="&rsbody("sbclass_id")&">�༭</a>&nbsp;&nbsp;<a href=sb_class.asp?action=delclass&id="&rsbody("sbclass_id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
	   Dwt.out "</Div></td></tr>"
	    			'����
			sqlz="SELECT * from sbclass where sbclass_zclass="&rsbody("sbclass_id")&" order by  sbclass_orderby aSC"& vbCrLf
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conn,1,1
			if rsz.eof and rsz.bof then 
			else
				dim xhz
				xhz=0
				do while not rsz.eof
				
					xhz=xhz+1
					if xhz mod 2 =1 then 
					  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"-"&xhz&"</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><b>"&rsz("sbclass_name")&"</b>&nbsp;</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&rsz("sbclass_zclass"))(0)&"-����</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsz("sbclass_isput")&"&nbsp;</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsz("sbclass_orderby")&"&nbsp;</Div></td>"
				   Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
				   Dwt.out "<a href=sb_class.asp?action=editjx&id="&rsz("sbclass_id")&">�༭��������</a>&nbsp;&nbsp;"
				   Dwt.out "<a href=sb_class.asp?action=edittable&id="&rsz("sbclass_id")&">�༭�������</a>&nbsp;&nbsp;"
				   if rsz("sbclass_zq") then 
				      dwt.out "<a href=sb_class.asp?action=sbclass_zq&id="&rsz("sbclass_id")&"&is=false onClick=""return confirm('ȷ���豸Ϊ����ʾ׼ȷ��');"">��ʾ׼ȷ</a>&nbsp;&nbsp;"
				   else
				      dwt.out "<a href=sb_class.asp?action=sbclass_zq&id="&rsz("sbclass_id")&"&is=true onClick=""return confirm('ȷ���豸Ϊ��ʾ׼ȷ��');"">����ʾ׼ȷ</a>&nbsp;&nbsp;"
    			   end if 	  
				   dwt.out "<a href=sb_class.asp?action=editclass&id="&rsz("sbclass_id")&">�༭</a>&nbsp;&nbsp;<a href=sb_class.asp?action=delclass&id="&rsz("sbclass_id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
				   Dwt.out "</Div></td></tr>"
					'����
					sqlzz="SELECT * from sbclass where sbclass_zclass="&rsz("sbclass_id")&" order by  sbclass_orderby aSC"& vbCrLf
					set rszz=server.createobject("adodb.recordset")
					rszz.open sqlzz,conn,1,1
					if rszz.eof and rszz.bof then 
					else
						dim xhzz
						xhzz=0
						do while not rszz.eof
						
					xhzz=xhzz+1
					if xhz mod 2 =1 then 
					  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
							Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"-"&xhz&"-"&xhzz&"</Div></td>"
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rszz("sbclass_name")&"&nbsp;</Div></td>"
							Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&rszz("sbclass_zclass"))(0)&"-����</Div></td>"
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rszz("sbclass_isput")&"&nbsp;</Div></td>"
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rszz("sbclass_orderby")&"&nbsp;</Div></td>"
						   Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
						   Dwt.out "<a href=sb_class.asp?action=editclass&id="&rszz("sbclass_id")&">�༭</a>&nbsp;&nbsp;<a href=sb_class.asp?action=delclass&id="&rszz("sbclass_id")&" onClick=""return confirm('ȷ��Ҫɾ����');"">ɾ��</a>"
						   Dwt.out "</Div></td></tr>"
						rszz.movenext
						loop
					end if 	
					rszz.close
					set rszz=nothing
				rsz.movenext
				loop
			end if 	
			rsz.close
			set rsz=nothing
		
    rsbody.movenext
    loop
     Dwt.out "</table></Div>"
end if 
  rsbody.close
  set rsbody=nothing
  'conn.close
  'set conn=nothing
  Dwt.out "</Div>"
end sub

Dwt.out "</body></html>"



'ȡ�ֶε�����
function sbtable_body(sbclass_id,sbtable_name)
dim sqlbody,rsbody
 sqlbody="SELECT sbtable_body from sbtable where sbtable_sbclassid="&sbclass_id&" and sbtable_name='"&sbtable_name&"'"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     sbtable_body= null
  else
     sbtable_body=rsbody("sbtable_body")
  end if
end function


'ȡ�ֶ�����˳��
function sbtable_orderby(sbclass_id,sbtable_name)
dim sqlbody,rsbody
 sqlbody="SELECT sbtable_orderby from sbtable where sbtable_sbclassid="&sbclass_id&" and sbtable_name='"&sbtable_name&"'"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     sbtable_orderby= 0 
  else
     sbtable_orderby=rsbody("sbtable_orderby")
  end if
end function
Call CloseConn
%>