<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<%dim sqlghname,rsghname,ghname
dim sqlbody,rsbody

sqlbody="SELECT * from ylbbody where id="&Trim(Request("id"))
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
   
  
	%>
	<html>
<head>
<title>�豸���������б�ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="22" class="topbg"><div align="center"><strong>�豸һ����</strong></div></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="tdbg2">
    <td> �� <a href="ylb.asp?action=1">��ӵ�ѹ����</a> | <a href="ylb.asp?action=3">ת����</a> | <a href="ylb.asp?action=4">���ڷ�����</a> | <a href="ylb.asp?action=5" >��ŷ�</a> | <a href="ylb.asp?action=6">�͵ص�����</a> | <a href="ylb.asp?action=7">ת��̽ͷ</a> | <a href="ylb.asp?action=8">����һ��Ԫ��</a> | <a href="ylb.asp?action=9">����һ��Ԫ��</a> | <a href="ylb.asp?action=10">����̽ͷ</a> | <a href="ylb.asp?action=11">����</a> | <a href="ylb.asp?action=12">�յ�</a> | <a href="ylb.asp?action=13">Ƥ����</a> | <a href="ylb.asp?action=14">���ڷ�</a> | <a href="ylb.asp?action=15">�綯ִ�л���</a></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="tdbg2">
    <td> �� άһ���� | ά������ | ά������</td>
  </tr>
</table>
<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>
  <tr>
    <td width="708" height='22'>�����ڵ�λ�ã�&nbsp;��������&nbsp;&gt;&gt;&nbsp;�豸һ����&nbsp;&gt;&gt;&nbsp;���ڷ�&gt;&gt;<%=rsbody("wh")%>��ϸ����</td>
  <td width='284' height='22' align='right'>
	  <select name='select' id='select4' onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
        <option value='' selected>��������ת����</option>
        <option value='ylb.asp?action=1'>άһ����</option>
        <option value='ylb.asp?action=2'>&nbsp;&nbsp;��&nbsp;bbbbbbbb</option>
        <option value='ylb.asp?action=2'>ά������</option>
        <option value='ylb.asp?action=2'>&nbsp;&nbsp;��&nbsp;bbbbbbbb</option>
      </select>
	  &nbsp;&nbsp;
	  <select name='JumpClass' id='JumpClass' onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
	       <option value='' selected>��������ת����</option>
		   <option value='ylb.asp?action=1'>��ӵ�ѹ����</option>
		   <option value='ylb.asp?action=2'>������</option>
		   <option value='ylb.asp?action=2'>&nbsp;&nbsp;��&nbsp;bbbbbbbb</option>
		   <option value='ylb.asp?action=3'>ת����</option>
		   <option value='ylb.asp?action=4'>���ڷ�����</option>
		   <option value='ylb.asp?action=5'>��ŷ�</option>
		   <option value='ylb.asp?action=6'>�͵ص�����</option>
		   <option value='ylb.asp?action=7'>ת��̽ͷ</option>
		   <option value='ylb.asp?action=8'>����һ��Ԫ��</option>
		   <option value='ylb.asp?action=9'>����һ��Ԫ��</option>
		   <option value='ylb.asp?action=10'>����̽ͷ</option>
  		   <option value='ylb.asp?action=11'>����</option>
		   <option value='ylb.asp?action=12'>�յ�</option>
		   <option value='ylb.asp?action=13'>Ƥ����</option>
		   <option value='ylb.asp?action=14'>���ڷ�</option>
		   <option value='ylb.asp?action=15'>�綯ִ�л���</option>
		   
		   
    </select>	</td> 
  </tr>
</table>&nbsp;
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>λ ��</strong></div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ͺ�</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>��������</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>CV����</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���ӷ�ʽ</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�¶�</strong></div></td>
    </tr>

        
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("wh")%>&nbsp;</div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("llname")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("ggxh")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_lltx")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_cv")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_ljfs")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("gyjz")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("czwd")%>&nbsp;</div></td>
    </tr>
    
		    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>���ѹ��</strong></div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>��P</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���ڷ��������</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���Ϲ��</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���ϲ���</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�з���</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�з����</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����</strong></div></td>
    </tr>
   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("czyl")%>&nbsp;</div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_dltp")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("ccbh")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_tlgg")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_tlcz")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("tjf_zfgg")%>&nbsp;</td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zfcz")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("shul")%>&nbsp;</div></td>
    </tr>
	
			    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>��������</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ִ�����������</strong>&nbsp;</div></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ִ�����ͺ�</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�г�ת��</strong>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����ѹ����</strong>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���ַ�ʽ</strong>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���÷�ʽ</strong>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�������쳧</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ִ�л������쳧</strong>&nbsp;</div></td>
    </tr>
   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("tjf_yxtl")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxqbh")%>&nbsp;</div></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxqxhgg")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_xczj")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_thyl")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_slfs")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("tjf_zyfs")%>&nbsp;</td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_fmcj")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxjgcj")%>&nbsp;</div></td>
    </tr>

	
</table>


<table width="100%"  border="0">
  <tr class="title">
    <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>�� ע</strong></div></td>
  </tr>
  <tr class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
    <td style="border-bottom-style: solid;border-width:1px" ><%=rsbody("whbeizhu")%>&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0">
  <tr class="title">
    <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center">���޼�¼&nbsp;������¼&nbsp;�༭������&nbsp;ɾ����λ����Ϣ</div></td>
  </tr>
</table>
	<%
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</body>
</html>
