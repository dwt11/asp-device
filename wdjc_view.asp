<!--#include file="conn.asp"-->
<!--#include file="inc/imgcode.asp"-->
<!--#include file="inc/function.asp"-->


<%
dim sql
dim rs
if request("name")="" or request("ssbz")="" or request("wz")="" then
response.write "�ó���ִ���˷Ƿ�����:)"
response.end
end if
%>
<html>
<link href='/DefaultSkin.css' rel='stylesheet' type='text/css'> 

<head>
<title><%=request("wz")%>-<%=request("name")%>-�¶�Ѳ���¼</title>

<link rel="stylesheet" href="../style/style.css" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
</TR>
</TBODY>
</TABLE>
</TD>
<TD></TD>
    </TR>
  </TBODY>
</TABLE>



  <table width="760" height="120" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td background="/images2006/logo.jpg">&nbsp;</td>
    </tr>
</table>
  
  <TABLE width=760 
  border=0 align="center" cellPadding=0 cellSpacing=0 class=table-left1-right2>
    <TBODY>
  <TR>
    <TD vAlign=bottom>
      <TABLE height=32 cellSpacing=0 cellPadding=0 width="100%" 
      background=images2006/topmenu-bg.gif border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE class=sxpta-font1 cellSpacing=0 cellPadding=0 width=750 
            border=0>
              <TBODY>
              <TR>
                <TD align=right width=41 
                background=images2006/topmenu-bg2.gif><IMG height=11 
                  src="images2006/d.gif" width=13 name=Image2></TD>
                <TD width=238 
                  background=images2006/topmenu-bg2.gif><div align="center"><A class=link2 
                  href="/index.asp">��ҳ</A></div></TD>
                <TD align=middle width=15><IMG height=32 
                  src="images2006/menuicon2.gif" width=15></TD>
                <TD align=middle width=341> ����</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>��������</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>������</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>ֵ���</TD>
                <TD align=middle width=115><IMG height=32 
                  src="images2006/menuicon.gif" width=15></TD>
                <TD align=middle width=341>֪ͨ</TD>
                </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
  <table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
    <tr>
      <td vAlign=top>
      <!--��վ�������´��뿪ʼ-->
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td class=main_title_1i>��ǰλ�ã�<a href="/" class=class>��Ϣ����ϵͳ��ҳ</a> &gt;&gt;&gt; <%=request("wz")%>-<%=request("name")%>-�¶�Ѳ���¼</td>
          </tr>
          
          <tr>
            <td  vAlign=top class=main_tdbg_575><br><div align="center"> <br>
             <%=request("wz")%>-<%=request("name")%> <br><hr width="80%" size="1">
              </div>
             
        <%
        response.Write "<br><table  border='1'  cellpadding='1' cellspacing='1' align=center>"
			  response.Write "		  <tr>"
			  response.Write "		    <td align=center>����</td>"
			  response.Write "		    <td align=center>λ��</td>"
			  response.Write "		    <td align=center>����</td>"
			  response.Write "		    <td align=center>�¶�</td>"
			  response.Write "	      </tr>"
		  sqljl="SELECT * from bb where ssbz="&request("ssbz")&" and name='"&request("name")&"' and wz='"&request("wz")&"' order by update desc"
		  set rsjl=server.createobject("adodb.recordset")
		  rsjl.open sqljl,connw,1,1
		  if rsjl.eof and rsjl.bof then 
		  response.Write "		  <tr>"
		  response.Write "		    <td colspan='5' align=center>δ��Ӽ�¼</td>"
		  response.Write "	      </tr>"
		  else
		  do while not rsjl.eof 
			  response.Write "		  <tr>"
			  response.Write "		    <td>"&rsjl("update")&"</td>"
			  response.Write "		    <td>"&rsjl("wz")&"</td>"
			  response.Write "		    <td>"&rsjl("name")&"</td>"
			  response.Write "		    <td>"&rsjl("ti")&"</td>"
			  response.Write "	      </tr>"
		
		
		  rsjl.movenext
		  loop
		  end if 
	
		  response.Write "</table><br>"
	

        
        %>     
             
             
             
             
             
             
             
             
             
             
             
             
             
             
             
              <table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td align="right"><div align="right"><font color="#990000"><br><br>��<a href="javascript:self.print()"><font color="#990000">��ӡ������</font></a>����<a href="javascript:window.close()"><font color="#990000">�رմ���</font></a>��</font> </div></td>
                </tr>
              </table>
			  

			  
			</td>
          </tr>
</table>
<TABLE width=760 border=0 align="center" cellPadding=0 cellSpacing=0 
background=images2006/bottom_back.gif>
  <TBODY>
  <TR>
    <TD class=sxpta-font2 align=middle height=24>�豸����ϵͳ</TD>
    <TD width=140 height=54 rowSpan=2><IMG height=54 src="images2006/bottom_r.gif" width=140   border=0></TD></TR>
  <TR>
    <TD class=sxpta-font2 align=middle height=30>
      <TABLE class=black cellSpacing=0 cellPadding=0 width=610 align=center 
      border=0>
        <TBODY>
        <TR>
          <TD width=170>
            </TD>
          <TD vAlign=bottom width=394 height=28>
</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>

</body>
</html>


