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
dim sqlpxst,rspxst,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel,wangong,rsdel2,sqldel2
url="anquangs_luoshi.asp"

keys=trim(request("keyword")) 

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>��Ϣ����ϵͳ����ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")
if request("wangong")="" then
wangong="ȫ��"
else
wangong=request("wangong")
end if

zr_danwei=request("zr_danwei")
ly=request("ly")




select case action
  case "add"
       'if truepagelevelh(session("groupid"),1,session("pagelevelid")) then 
	   call add
  case "saveadd"
    call saveadd
  case "edit"
	'if truepagelevelh(session("groupid"),2,session("pagelevelid")) then 
	call edit
  case "saveedit"
    call saveedit
  case "del"
    'if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
	call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then
	 call main
end select	

sub add()
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�� �� �� ��</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='anquangs_luoshi.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�̰�����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_title>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >���Ĵ�ʩ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_zgcs>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&session("username1")&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&date()&"'  disabled='disabled' >"& vbCrLf
	dwt.out "<input name='pxst_date' type='hidden' value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >���ε�λ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_ren>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ҫ���깤ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='yaoqiu_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	dwt.out"</div> "& vbCrLf  
	
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from anquangs" 
      rsadd.open sqladd,connaq,1,3
      rsadd.addnew
   	  
	  rsadd("pxst_title")=Trim(Request("pxst_title"))
      rsadd("pxst_body")=Trim(request("pxst_body"))
      rsadd("pxst_zgcs")=Trim(request("pxst_zgcs"))
      rsadd("pxst_date")=request("pxst_date")
     ' rsadd("pxst_class")=request("anquangs_class")
      rsadd("userid")=session("userid")
      rsadd("huiyi_date")=request("txdate")
      rsadd("edit_date")=request("pxst_date")
      rsadd("zr_danwei")=request("zr_danwei")
      rsadd("yaoqiu_date")=request("yaoqiu_date")
      rsadd("zr_ren")=request("zr_ren")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='anquangs_luoshi.asp';</Script>"
end sub

sub edit()
     '�༭
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from anquangs where id="&id
   rsedit.open sqledit,connaq,1,1
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>�༭��ȫ�̰�����</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='anquangs_luoshi.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�̰�����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_title value='"&rsedit("pxst_title")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >���Ĵ�ʩ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_zgcs value='"&rsedit("pxst_zgcs")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("huiyi_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&usernameh(rsedit("userid"))&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>����ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&rsedit("pxst_date")&"'  disabled='disabled' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >���ε�λ:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_danwei value='"&rsedit("zr_danwei")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=zr_ren value='"&rsedit("zr_ren")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>Ҫ���깤ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='yaoqiu_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("yaoqiu_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ʵ���깤ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
    dwt.out"<input name='wangong_date' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("wangong_date")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�깤���:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	    dwt.out"<input name='wanggong' type='radio' value='true'"
	if rsedit("isno")=true then dwt.out "checked"
	dwt.out" />���깤 <input name='wanggong' type='radio' value='false' "
	if rsedit("isno")=false then dwt.out "checked"
	dwt.out" />δ�깤"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >Ч������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 375px' name=pxst_estimation value='"&rsedit("pxst_estimation")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >�� �� ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=pxst_estimator value='"&rsedit("pxst_estimator")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
		

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' ���� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	dwt.out"</div> "& vbCrLf  


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from anquangs where ID="&ReplaceBadChar(Trim(request("ID")))
rsedit.open sqledit,connaq,1,3
rsedit("pxst_title")=ReplaceBadChar(Trim(Request("pxst_title")))
rsedit("pxst_body")=Trim(request("pxst_body"))
rsedit("pxst_zgcs")=Trim(request("pxst_zgcs"))
'rsedit("pxst_class")=Trim(request("anquangs_class"))
rsedit("edit_date")=date()
rsedit("huiyi_date")=request("txdate")
rsedit("yaoqiu_date")=request("yaoqiu_date")
rsedit("wangong_date")=request("wangong_date")
rsedit("zr_danwei")=request("zr_danwei")
rsedit("zr_ren")=request("zr_ren")
rsedit("isno")=trim(request("wanggong"))
rsedit("pxst_estimation")=Trim(request("pxst_estimation"))
rsedit("pxst_estimator")=Trim(request("pxst_estimator"))
rsedit.update
rsedit.close
	  dwt.out"<Script Language=Javascript>location.href='anquangs_luoshi.asp';</Script>"
end sub


sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>��ȫ�̰幫ʾ"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
'	search()
'dim sqlcj,rscj
dwt.out "<div class='x-toolbar'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='anquangs_luoshi.asp'>" & vbCrLf
dwt.out "<a href=""anquangs_luoshi.asp?action=add"">�������</a>&nbsp;&nbsp;����������" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
		 dwt.out "<select name='wangong'>" 
	dwt.out "<option value='ȫ��'"
	if wangong="ȫ��" then dwt.out "selected"
	dwt.out ">ȫ��</option>" 
	dwt.out "<option value='�깤'"
	if wangong="�깤" then dwt.out "selected"
	dwt.out ">�깤</option>" 
	dwt.out "<option value='δ�깤'"
	if wangong="δ�깤" then dwt.out "selected"
	dwt.out ">δ�깤</option>" 
    dwt.out "</select>" 


dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
dwt.out "</form></div>" & vbCrLf

'	sqlpxst="SELECT * from anquangs " 
'	if request("classid")<>"" then sqlpxst=sqlpxst&" where pxst_class="&request("classid")&" and pxst_body like '%"&keys&"%'"
'	sqlpxst=sqlpxst&" ORDER BY id DESC"
	
		sqlpxst="SELECT * from anquangs where 1=1 "

	
	

	if wangong="�깤" then 
	sqlpxst=sqlpxst& " and  isno=true  "
	end if
	if wangong="δ�깤" then 
	sqlpxst=sqlpxst& " and  isno=false  "
	end if

	if keys<>"" then 
	sqlpxst=sqlpxst& "  and  pxst_body like '%" &keys& "%'  "
	end if 
	
		sqlpxst=sqlpxst&"  ORDER BY id DESC"
'dwt.out sqlpxst
	set rspxst=server.createobject("adodb.recordset")
	rspxst.open sqlpxst,connaq,1,1
	if rspxst.eof and rspxst.bof then 
	message("δ�ҵ��������")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td' ><DIV class='x-grid-hd-text'>���</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>���ʱ��</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>Ҫ��ʱ��</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>�깤ʱ��</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>��  ��  ��  ��</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>��  ��  ��  ʩ</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>���ε�λ</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>������</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>������</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>������</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>Ч������</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>������</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>ѡ��</div></td>"
		dwt.out "    </tr>"
           record=rspxst.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxst.PageSize = Cint(PgSz) 
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
           rspxst.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxst.PageSize
           do while not rspxst.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("huiyi_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("yaoqiu_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("wangong_date")&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"
				 
				 
				 
				 
				 dwt.out"<a href=anquangs_view.asp?id="&rspxst("id")&" target=_blank>"
				 
				 	if rspxst("isno")=true then

				dwt.out rspxst("pxst_title")
				else
				dwt.out "<strong><div style=' color: #F00;'>"&rspxst("pxst_title")&"</div></strong>"
				end if  
				 
				
				 
				 dwt.out"</a></td>"
         dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("pxst_zgcs")&"&nbsp;</td>"
         dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("zr_danwei")&"&nbsp;</td>"
          dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("zr_ren")&"&nbsp;</td>"
           
				 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 if isnull(rspxst("userid")) then 
				   dwt.out rspxst("pxst_zz")
				 else
				   dwt.out usernameh(rspxst("userid")) 
				 end if   
				 dwt.out"&nbsp;</div></td>"
				 
'                 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_date")&"</div></td>"
				 
	if rspxst("isno")=true then
		dwt.out "<td  style=""border-bottom-style: solid;border-width:1px;white-space:nowrap"" ><div align=""center"">���</div></td>"
	else
		dwt.out "<td  style=""border-bottom-style: solid;border-width:1px;white-space:nowrap"" ><div align=""center"">δ���</div></td>"
	end if
	
         dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("pxst_estimation")&"&nbsp;</td>"
         dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&rspxst("pxst_estimator")&"&nbsp;</td>"
	
	
				 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 
				 			if session("levelclass")<>0 then 
				dim rsqxtb_fk,sqlqxtb_fk
				set rsqxtb_fk=server.createobject("adodb.recordset")
				sqlqxtb_fk="select * from anquangs_fk where huiyiluoshi_fk_sscj="&session("levelclass")&" and huiyiluoshi_id="&rspxst("id")
				rsqxtb_fk.open sqlqxtb_fk,connaq,1,1
				if rsqxtb_fk.eof and rsqxtb_fk.bof then 
					dwt.out  "<a href='anquangs_fk.asp?action=add&huiyiluoshi_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"'>��ӷ���</a>&nbsp;"
				else
					dwt.out  "<a href='anquangs_fk.asp?action=edit&huiyiluoshi_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"'>�༭����</a>&nbsp;"
					if session("level")=0 then dwt.out  "<a href='anquangs_fk.asp?action=del&qxtb_fk_sscj="&session("levelclass")&"&huiyiluoshi_id="&rspxst("id")&"' onClick=""return confirm('ȷ��Ҫɾ���˷�����');"">ɾ������</a>"
				end if 
				rsqxtb_fk.close
				set rsqxtb_fk=nothing
			end if 

				 
				 
				 if session("level")=0 or session("levelclass")=9 or rspxst("userid")=session("userid") then
				  dwt.out "<a href='anquangs_luoshi.asp?action=edit&ID="&rspxst("id")&"'>�༭</a>"
				  dwt.out "&nbsp;<a href='anquangs_luoshi.asp?action=del&ID="&rspxst("id")&"' onClick=""return confirm('ȷ��Ҫɾ������Ŀ��');"">ɾ��</a>"
				 end if 			'call editdel(rspxst("id"),rspxst("sscj"),"anquangs_luoshi.asp?action=edit&id=","anquangs_luoshi.asp?action=del&id=")
				 dwt.out "&nbsp; </div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rspxst.movenext
          loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
	rspxst.close
	set rspxst=nothing
	conn.close
	set conn=nothing
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from anquangs_fk where huiyiluoshi_id="&id
rsdel.open sqldel,connaq,1,3
set rsdel=nothing 
 
set rsdel2=server.createobject("adodb.recordset")
sqldel2="delete * from anquangs where id="&id
rsdel2.open sqldel2,connaq,1,3
dwt.out"<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel2=nothing  

end sub


Function class_name(class_id)
    dim sqlcj,rscj
'dim class_id

	  sqlcj="SELECT * from anquangs_class where id="&class_id
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,connpxjhzj,1,1
    if rscj.eof then 
		class_name="δ�༭"
	else
	do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    class_name=rscj("class_name")
		rscj.movenext
	loop
	end if 
	rscj.close
	set rscj=nothing
end Function
dwt.out "</body></html>"


sub search()
dim sqlcj,rscj
dwt.out "<div class='x-toolbar'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='anquangs_luoshi.asp'>" & vbCrLf
dwt.out "<a href=""anquangs_luoshi.asp?action=add"">�������</a>&nbsp;&nbsp;����������" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
dwt.out "</form></div>" & vbCrLf
end sub

Call CloseConn
%>