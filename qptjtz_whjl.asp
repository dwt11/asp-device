<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<%
'dim sqlqptjtz,rsqptjtz,title,record,pgsz,total,page,start,rowcount,xh,url,ii
'dim rsadd,sqladd,qptjtzid,rsedit,sqledit,scontent,rsdel,sqldel,sscj,tyzk,id,sscjh,qptjtzwh,sql,rs,czjg
dim qptjtzid,qptjtzbh,sql,rs,sqlqptjtz,rsqptjtz,rsadd,sqladd,rsedit,sqledit
dim record,pgsz,total,page,start,rowcount,url,ii
dim czjg,id,rsdel,sqldel,onnumb,sqld,rsd,rscj,sqlcj
qptjtzid=Trim(Request("qptjtzid"))

dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>��Ϣ����ϵͳ��ƿ̨�ʹ���ҳ</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out "if(document.form1.qptjtz_sscj.value==''){" & vbCrLf
dwt.out "alert('��ѡ��ʹ�õ�λ��');" & vbCrLf
dwt.out "document.form1.qptjtz_sscj.focus();" & vbCrLf
dwt.out "return false;" & vbCrLf
dwt.out "}" & vbCrLf

dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf

dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0' onload='Javascript:document.getElementById(""Submit"").disabled=true;'>"& vbCrLf
if Request("action")="add" then call add
if Request("action")="add1" then call add1
if Request("action")="saveadd" then call saveadd
if Request("action")="saveadd1" then call saveadd1
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del
if request("action")="" then call main 
sub add1()
	dwt.out "<br><br><br><form method='post' action='qptjtz_whjl.asp' name='form6' >"
	dwt.out "<table width='20%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out "<tr class='title'><td height='22' colspan='2'>"
	dwt.out "<div align='center'><strong>Ͷ��״̬</strong></div></td>    </tr>"
	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='qptjtz_tyqk' size='1'>"
	dwt.out"<option value='2'>����</option>"
	dwt.out"<option value='3'>�˿�</option>"
    dwt.out"</select></td></tr>"

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd1'> <input name='qptjtz_whjl_qptjtzid' type='hidden'  value='"&Trim(Request("qptjtzid"))&"'>    <input  type='submit' name='Submit1' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	
sub saveadd1()
    dim aa,bb
	  qptjtzid=Trim(request("qptjtz_whjl_qptjtzid"))
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from qptjtz_whjl" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
	  if request("qptjtz_tyqk")=2 then
				aa="��ƿ������ڣ����˿�"
	  else
				aa="�˿�"
	  end if
      rsadd("whyy")=aa
      rsadd("whsj")=date()
      rsadd("whjg")=request("qptjtz_tyqk")
	  rsadd("qpjxid")=qptjtzid
      rsadd.update
      rsadd.close
      set rsadd=nothing

	  '����
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from qptjtz where qptzid="&qptjtzid
      rsedit.open sqledit,connjg,1,3
	  if request("qptjtz_tyqk")=3 then
	  rsedit("tcdata")=date()
	  else
	  rsedit("tcdata")=null
	  end if
      rsedit("updata")=now()
	  rsedit("tyqk")=request("qptjtz_tyqk")		  
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  dwt.savesl "��ƿά����¼","�½�",Connjg.Execute("SELECT bh FROM qptjtz WHERE qptzid="&trim(request("qptjtz_whjl_qptjtzid"))&"")(0) 
	  dwt.out "<Script Language=Javascript>location.href='qptjtz.asp';</Script>"
end sub

sub add()
  Dwt.out"<script type=""text/javascript"" src=""js/checkbh.js""></script>"&vbcrlf
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from qptjtz where qptzid="&qptjtzid
   rsedit.open sqledit,connjg,1,1

	qptjtzbh=Connjg.Execute("SELECT bh FROM qptjtz WHERE qptzid="&qptjtzid)(0)
	dwt.out "<br><br><br><form method='post' action='qptjtz_whjl.asp' name='form1'  onsubmit='javascript:return checkadd();' >"
	dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out "<tr class='title'><td height='22' colspan='2'>"
	dwt.out "<div align='center'><strong>�����ƿ̨��  "&rsedit("bh")&"  ������¼</strong></div></td>    </tr>"
   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>ʹ�õ�λ�� </strong></td>"      
   dwt.out"<td  class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='qptjtz_sscj' size='1'>"
    dwt.out"<option >��ѡ��ʹ�õ�λ</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
    dwt.out"<input name='qptjtz_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    dwt.out"<input name='qptjtz_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
 end if 
 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_bqname' value='"&rsedit("bqname")&"' ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ƿ�����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qptj' value='"&rsedit("qptj")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ƿѹ����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyl' value='"&rsedit("qpyl")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ɷݺ�����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpcf' value='"&rsedit("qpcf")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��Ʒ��ţ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_ypbh' value='"&rsedit("ypbh")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyq' value='"&rsedit("qpyq")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_yt' ></td></tr> "
	 
	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='qptjtz_tyqk' size='1'>"
	dwt.out"<option value='1'>����</option>"
	dwt.out"<option value='2'>����</option>"
	dwt.out"<option value='3'>�˿�</option>"
    dwt.out"</select></td></tr>"

	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>����ԭ��</strong></td>"
	dwt.out "<td width='88%' class='tdbg'><input name='qptjtz_whjl_whyy' type='text'></td>    </tr>   "
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ֵ����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='qptjtz_scdata' style='WIDTH: 175px'  value="&date()&"  onClick='new Calendar(0).show(this)' readOnly >"
	dwt.out"</td></tr>  "   
	
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��Ч�ڣ�</strong></td> "
	Dwt.Out"<td width='80%' class='tdbg'>"	
	dwt.out outdatadict2 ("qptjtz_yxq","��Ч��",onnumb,rsedit("yxq"))

    dwt.out "</td></tr>"
	 
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "

	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='qptjtz_lydata' type='text' onClick='new Calendar(0).show(this)' readOnly  value="&date()&" >"
	dwt.out "</td></tr>"& vbCrLf
	
   	 dwt.out"<tr><td width='12%' align='right' class='tdbg'><strong>��������</strong></td>"      
	 Dwt.out "<td width='88%' class='tdbg'><input name='qptjtz_dqdata' type='text' id='input6' onFocus='return addrdata()'/>&nbsp;<span >����Զ�����</span>"

	dwt.out"</td></tr>  " 
	  
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ŵص㣺</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_cfdd' value='"&rsedit("cfdd")&"'  ></td></tr> "
  	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ע��</strong></td> "
	dwt.out "<td><input name='fdbw_whjl_body' type='text'></td></tr>"


	       rsedit.close
       set rsedit=nothing

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd'> <input name='qptjtz_whjl_qptjtzid' type='hidden'  value='"&Trim(Request("qptjtzid"))&"'>    <input  type='submit' name='Submit' id='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	

sub saveadd()    
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from qptjtz_whjl" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("whyy")=Trim(Request("qptjtz_whjl_whyy"))
      rsadd("bz")=request("qptjtz_whjl_body")
      rsadd("whsj")=Trim(request("qptjtz_lydata"))
      rsadd("whjg")=request("qptjtz_tyqk")
	  qptjtzid=Trim(request("qptjtz_whjl_qptjtzid"))
      rsadd("qpjxid")=qptjtzid
	  
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from qptjtz where qptzid="&qptjtzid
      rsedit.open sqledit,connjg,1,3
      rsedit("updata")=now()
	  rsedit("scdata")=request("qptjtz_scdata")

	  if request("qptjtz_tyqk")=3 then
	  rsedit("tcdata")=Trim(request("qptjtz_lydata"))
	  else
	  rsedit("ghdata")=Trim(request("qptjtz_lydata"))
	  rsedit("tcdata")=null
      end if
	  rsedit("yt")=Trim(request("qptjtz_yt"))
	  rsedit("yxq")=Trim(request("qptjtz_yxq"))
	  rsedit("sscj")=Trim(Request("qptjtz_sscj"))
      rsedit("bqname")=request("qptjtz_bqname") 
      rsedit("qptj")=Trim(request("qptjtz_qptj"))
      rsedit("qpyl")=request("qptjtz_qpyl")
      rsedit("qpcf")=request("qptjtz_qpcf")
      rsedit("qpyq")=request("qptjtz_qpyq")
      rsedit("ypbh")=request("qptjtz_ypbh")
	  rsedit("cfdd")=request("qptjtz_cfdd")
      rsedit("dqdata")=request("qptjtz_dqdata")
	  rsedit("tyqk")=request("qptjtz_tyqk")
	  rsedit("bz")=request("qptjtz_whjl_body")
		  
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	
	  dwt.savesl "��ƿά����¼","�½�",Connjg.Execute("SELECT bh FROM qptjtz WHERE qptzid="&trim(request("qptjtz_whjl_qptjtzid"))&"")(0) 
	  dwt.out "<Script Language=Javascript>location.href='qptjtz_whjl.asp?qptjtzid="&qptjtzid&"';</Script>"
end sub


sub saveedit()    
	  '����
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from qptjtz_whjl where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connjg,1,3
      rsedit("whyy")=Trim(Request("qptjtz_whjl_whyy"))
      rsedit("bz")=request("qptjtz_whjl_body")
      rsedit("whsj")=Trim(request("qptjtz_lydata"))
	  rsedit("whjg")=request("qptjtz_tyqk")
	  
	  
	  set rs=server.createobject("adodb.recordset")
      sql="select * from qptjtz where qptzid="&rsedit("qpjxid")
      rs.open sql,connjg,1,3
	  rs("updata")=now()
	  rs("scdata")=request("qptjtz_scdata")

	  if request("qptjtz_tyqk")=3 then
	  rs("tcdata")=Trim(request("qptjtz_lydata"))
	  else
	  rs("ghdata")=Trim(request("qptjtz_lydata"))
	  rs("tcdata")=null
      end if
	  
	  rs("yt")=Trim(request("qptjtz_yt"))
	  rs("yxq")=Trim(request("qptjtz_yxq"))
	  rs("sscj")=Trim(Request("qptjtz_sscj"))
      rs("bqname")=request("qptjtz_bqname") 
      rs("qptj")=Trim(request("qptjtz_qptj"))
      rs("qpyl")=request("qptjtz_qpyl")
      rs("qpcf")=request("qptjtz_qpcf")
      rs("qpyq")=request("qptjtz_qpyq")
      rs("ypbh")=request("qptjtz_ypbh")
	  rs("cfdd")=request("qptjtz_cfdd")
      rs("dqdata")=request("qptjtz_dqdata")
	  rs("tyqk")=request("qptjtz_tyqk")	
	  rs("bz")=request("qptjtz_whjl_body")
	  
	  rs.update
      rs.close
      set rs=nothing

	  dwt.savesl "��ƿά����¼","�༭",Connjg.Execute("SELECT bh FROM qptjtz WHERE qptzid="&rsedit("qpjxid")&"")(0) 

      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out "<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
 	sqledit="select * from qptjtz_whjl where ID="&id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,connjg,1,1
    dwt.savesl "��ƿά����¼","ɾ��",Connjg.Execute("SELECT bh FROM qptjtz WHERE qptzid="&rsedit("qpjxid")&"")(0) 
	rsedit.close
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from qptjtz_whjl where id="&id
  rsdel.open sqldel,connjg,1,3
  dwt.out "<Script Language=Javascript>history.back()</Script>"
set rsdel=nothing  

end sub


sub edit()
  Dwt.out"<script type=""text/javascript"" src=""js/checkbh.js""></script>"&vbcrlf

  sql="SELECT * from qptjtz where qptzid="&qptjtzid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
qptjtzbh=rs("bh")
rs.close
   dwt.out "<br><br><br><form method='post' action='qptjtz_whjl.asp' name='form1'  onsubmit='javascript:return checkadd();' >"
   dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>�༭��ƿ̨��  "&qptjtzbh&"  ������¼</strong></div></td>    </tr>"
     
   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>ʹ�õ�λ�� </strong></td>"      
   dwt.out"<td  class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='qptjtz_sscj' size='1'>"
    dwt.out"<option >��ѡ��ʹ�õ�λ</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
    dwt.out"<input name='qptjtz_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    dwt.out"<input name='qptjtz_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
 end if 
 
    set rs=server.createobject("adodb.recordset")
   sql="select * from qptjtz where qptzid="&qptjtzid
   rs.open sql,connjg,1,3

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�������ƣ�</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_bqname' value='"&rs("bqname")&"' ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ƿ�����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qptj' value='"&rs("qptj")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ƿѹ����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyl' value='"&rs("qpyl")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ɷݺ�����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpcf' value='"&rs("qpcf")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;����</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyq' value='"&rs("qpyq")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;;��</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_yt' value='"&rs("yt")&"'  ></td></tr> "
	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>״̬��</strong></td>"
	dwt.out"<td><select name='qptjtz_tyqk' size='1'>"
	dwt.out"<option value='1'>����</option>"
	dwt.out"<option value='2'>����</option>"
	dwt.out"<option value='3'>�˿�</option>"
    dwt.out"</select></td></tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ŵص㣺</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_cfdd' value='"&rs("cfdd")&"'  ></td></tr> "
	

	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��Ч�ڣ�</strong></td> "
	Dwt.Out"<td width='80%' class='tdbg'>"	
	dwt.out outdatadict2 ("qptjtz_yxq","��Ч��",onnumb,rs("yxq"))

    dwt.out "</td></tr>"
 	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��ֵ����</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='qptjtz_scdata' style='WIDTH: 175px'  value='"&rs("scdata")&"'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"</td></tr>  "   

	 
	rs.update
      rs.close
      set rs=nothing 
	 
 id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from qptjtz_whjl where id="&id
   rsedit.open sqledit,connjg,1,3

	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>����ԭ��</strong></td>"
	dwt.out "<td width='20%' class='tdbg'><input name='qptjtz_whjl_whyy' type='text'value='"&rsedit("whyy")&"'></td>    </tr>   "

	
	 
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺</strong></td> "

	dwt.out "<td width='88%' class='tdbg'>"
	dwt.out "<input name='qptjtz_lydata' type='text' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("whsj")&"'>"
	dwt.out "</td></tr>"& vbCrLf
	
   	 dwt.out"<tr><td width='12%' align='right' class='tdbg'><strong>��������</strong></td>"      
	 Dwt.out "<td width='88%' class='tdbg'><input name='qptjtz_dqdata' type='text' id='input6' onFocus='return addrdata()'/>&nbsp;<span >����Զ�����</span>"

	dwt.out"</td></tr>  "   
	
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	dwt.out "<strong>��ע��</strong></td>"
	dwt.out "<td width='60%' class='tdbg'><input name='qptjtz_whjl_body' type='text'value='"&rsedit("bz")&"'></td>    </tr>   "
	
	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveedit'><input type='hidden' name='id' value='"&id&"'> <input  type='submit' name='Submit'  id='Submit' value=' ��  �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
      rsedit.update
      rsedit.close
      set rsedit=nothing

end sub


sub main()
dim lb,brxx
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  "<tr class='topbg'>"& vbCrLf
dwt.out  "<td height='22' colspan='2' align='center'><strong>��ƿ̨�ʣ�������¼</strong></td>"& vbCrLf
dwt.out  "</tr>"& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "<td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
dwt.out  "<td height='30'><a href=""qptjtz.asp"">��ƿ̨����ҳ</a>&nbsp;|&nbsp;<a href=""qptjtz.asp?action=add"">�����ƿ̨��</a>"
dwt.out  "</td>"& vbCrLf
dwt.out  "  </tr>"& vbCrLf
dwt.out  "</table>"& vbCrLf

sql="SELECT * from qptjtz where qptzid="&qptjtzid
set rs=server.createobject("adodb.recordset")
rs.open sql,connjg,1,1
if session("levelclass")=rs("sscj") or session("level")=0 then 
	dwt.out  "<a href='qptjtz_whjl.asp?action=add&qptjtzbh="&qptjtzbh&"&qptjtzid="&qptjtzid&"'>�����ƿ̨��<font color='#ff0000'>"&rs("bh")&"</font>������¼</a>"
 else
    dwt.out  "&nbsp;"
 end if 
 dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
dwt.out  "<tr class=""title"">"  & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center""><strong>ʹ�õ�λ</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ƿ�ɷ�</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ƿ����</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ƿѹ��</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ƿ���</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ŵص�</strong></div></td>" & vbCrLf
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ʱ��</strong></div></td>" & vbCrLf

dwt.out  "    </tr>" & vbCrLf
                 dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&sscjh_d(rs("sscj"))&"</div></td>" & vbCrLf

                dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("bqname")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("qpcf")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("qpyq")&"&nbsp;</div></td>" & vbCrLf				
	            dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("qpyl")&"&nbsp;</div></td>" & vbCrLf
	            dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("qptj")&"&nbsp;</div></td>" & vbCrLf
		        dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("cfdd")&"&nbsp;</div></td>" & vbCrLf
		        dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rs("dqdata")&"&nbsp;</div></td>" & vbCrLf

 dwt.out  " </tr></table>"
rs.close
set rs=nothing


dwt.out  "<div align='center'>ά����¼</div>"
sqlqptjtz="SELECT * from qptjtz_whjl where qpjxid="&qptjtzid&" ORDER BY id DESC"
set rsqptjtz=server.createobject("adodb.recordset")
rsqptjtz.open sqlqptjtz,connjg,1,1
if rsqptjtz.eof and rsqptjtz.bof then 
dwt.out  "<p align='center'>δ�����ƿ̨��<font color='#ff0000'>"&qptjtzbh&"</font>������¼</p>" 
else
dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
dwt.out  "<tr class=""title"">" 
dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>���</strong></div></td>"
dwt.out  "      <td width=""40%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ԭ��</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ʱ��</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ά�����</strong></div></td>"
dwt.out  "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ע</strong></div></td>"

dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"

dwt.out  "    </tr>"
           record=rsqptjtz.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsqptjtz.PageSize = Cint(PgSz) 
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
           rsqptjtz.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsqptjtz.PageSize
           do while not rsqptjtz.eof and rowcount>0
      		
			dim tyqk
				select case rsqptjtz("whjg")
			  case 1
				 tyqk="<span style='color:#006600'>����</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>����</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>�˿�</span>"
			end select	 
                 dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
          dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&rsqptjtz("id")&"</div></td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rsqptjtz("whyy")&"&nbsp;</td>"
                dwt.out  "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("whsj")&"&nbsp;</div></td>"
        		dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&tyqk&"&nbsp;</div></td>"
				dwt.out  "<td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("bz")&"&nbsp;</div></td>"

                dwt.out  "<td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=center>"
				sql="SELECT * from qptjtz where qptzid="&qptjtzid
                set rs=server.createobject("adodb.recordset")
                rs.open sql,connjg,1,1
				call editdel(rsqptjtz("id"),rs("sscj"),"qptjtz_whjl.asp?action=edit&qptjtzid="&qptjtzid&"&id=","qptjtz_whjl.asp?action=del&id=")
				rs.close
                set rs=nothing

                dwt.out  "</div></td></tr>"

                 RowCount=RowCount-1
          rsqptjtz.movenext
          loop
        dwt.out  "</table>"
       call showpage1(page,url,total,record,PgSz)
       end if
       rsqptjtz.close
       set rsqptjtz=nothing
        connjg.close
        set connjg=nothing

end sub
dwt.out  "</body></html>"
Call Closeconn
%>