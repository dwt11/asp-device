<!--#include file="conn.asp"-->
<!--#include file="inc/imgcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<%

vip=request.servervariables("http_x_forwarded")
if vip="" then
vip=request.servervariables("remote_addr")

end if
dim sql
dim rs
if request("id")<>"" then
	set rsnews=server.createobject("adodb.recordset")
	sqlnews="select * from csyy_body where id="&request("id")
	rsnews.open sqlnews,conncs,1,1
	title=rsnews("news_title")
end if
%>
<html>
<link href='css/index.css' rel='stylesheet' type='text/css'> 

<head>
<title><%=title%>-��Ϣ����ϵͳ</title>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >

<%
action=request("action")

select case action
  case "savepost"
        call savepost
  case "addpost"
    call addpost
  case "del"
	if session("groupid")=3 then call del
  case ""
	 call main
end select	
'if request("action")="" then call main
'if request("action")="savepost" then call savepost
'if request("action")="addpost" then call addpost
'if 
sub savepost()    
	  '����
      if Request("body")<>"" then
		  dwt.savesl "�������Իظ�","���",ReplaceBadChar(Trim(Request("body")))
		  set rsadd=server.createobject("adodb.recordset")
		  sqladd="select * from csyy_re" 
		  rsadd.open sqladd,conncsyy,1,3
		  rsadd.addnew
		  rsadd("body")=ReplaceBadChar(Trim(Request("body")))
		  rsadd("news_id")=request("news_id")
                 RSADD("IP")=vip
		  rsadd.update
		  rsadd.close
		  set rsadd=nothing
      end if
	  dwt.out "<Script Language=Javascript>location.href='news_csyy_view.asp?id="&request("news_id")&"';</Script>"
end sub





sub main()
'	set rsnews=server.createobject("adodb.recordset")
'	sqlnews="select * from xzgl_news where id="&request("id")
'	rsnews.open sqlnews,conna,1,1


%>
	<!--#include file="index_t.asp"-->
<!--��վ�������´��뿪ʼ-->


<table width="760" border=0 align="center" cellpadding=0 cellspacing=0>
  <tr>
    <td class=main_title_1i>��ǰλ�ã�<a href="/" class=class>��Ϣ����ϵͳ��ҳ</a>&gt;&gt;&gt; <a href="news_csyy.asp?CLASSid=1">��������</a> &gt;&gt;&gt; <%=rsnews("news_title")%></td>
  </tr>
  <tr>
    <td  valign=top class=main_tdbg_575><br>
      <div align="center"><font color="#05006c" size=larger><%=rsnews("news_title")%></font> <br>
        <br>
        <hr width="80%" size="1">
      </div>
      <table width="80%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <%if rsnews("news_zz")<>"" then 
				           news_zz=rsnews("news_zz")
					else
					    news_zz=rsnews("user_id")	
					end if    
				  %>
          <td align="center">����ʱ�䣺<font color="#990000"><%=rsnews("news_date")%></font> ������<font color="#990000"><%=news_zz%> </font></td>
        </tr>
        <tr>
          <td><br>
            <%
   response.write imgCode(rsnews("news_body"))
 	Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from  csyy_body where id="&request("id")
    rs.Open sql, Conncs, 1, 3
	rs("view_numb") =rs("view_numb")+1      
	rs.Update
	rs.Close
	sql="SELECT isre FROM csyy_class WHERE id="&rsnews("news_class")
	isre=conncs.Execute(sql)(0)

    if isre then 
		dwt.out "<br/><br/><br/>��ǰ1¥&nbsp;&nbsp;&nbsp;&nbsp;<a href=news_csyy_view.asp?action=addpost&news_id="&rsnews("id")&"><span style=""border-bottom-style: solid;border-width:2px;color:#660066"">�ظ�����</span></a><br/> <hr width=100% size=1><br/><br/><br/>"
		post=1
		sqlpost="SELECT * FROM csyy_re WHERE news_id="&rsnews("id")&" order by DATE "
		set rspost=server.createobject("adodb.recordset")
		rspost.open sqlpost,conncs,1,1
		if rspost.eof and rspost.bof then 
		else
		   do while not rspost.eof
			post=post+1
			dwt.out rspost("body")
			dwt.out "<br/><br/><br/>��ǰ"&post&"¥&nbsp;&nbsp;&nbsp;"&rspost("date")
		 if  session("groupid")=3  then 
				 dwt.out "&nbsp;&nbsp;&nbsp;<a href='news_csyy_view.asp?action=del&ID3="&rspost("id")&"' onClick=""return confirm('ȷ��Ҫɾ����������');""><span style=""border-bottom-style: solid;border-width:2px;color:#660066"">ɾ��</span></a>"
				 end if
			dwt.out "<br/><hr width=100% size=1><br/><br/><br/>"
		  rspost.movenext
		  loop
        end if
	end if  
    
%>
           </td>
        </tr>
        <tr>
          <td align="right"><div align="right"><font color="#990000"><br>
              <br>
              ��<a href="javascript:self.print()"><font color="#990000">��ӡ������</font></a>����<a href="javascript:window.close()"><font color="#990000">�رմ���</font></a>��</font> </div>
  </td>      </tr>
      </table>
<%
 if isre then 
    dwt.out "<div align=left> <hr width=100% size=1>"
	dwt.out "<form method='post' action='news_csyy_view.asp' name='form1' align=center>"
	dwt.out "   <textarea name='body' cols='40' rows='10'></textarea><br/>"	
	dwt.out"<input name='action' type='hidden' value='savepost'> <input name='news_id' type='hidden' value="&rsnews("id")&">      <input  type='submit' name='Submit' value=' �ظ� ' style='cursor:hand;'>"& vbCrLf
    dwt.out "</form></div>	"
end if   

rsnews.close
set rsnews=nothing

  
%>	  
    </td>
  </tr>
  <tr>
    <td ></td>
  </tr>
</table>
<!--������������-->
<!--����Ƶ����ʾ����-->





<table width=760 border=0 align="center" cellpadding=0 cellspacing=0 
background=images2006/bottom_back.gif>
  <tbody>
    <tr>
      <td class=sxpta-font2 align=middle height=24>�豸����ϵͳ</td>
      <td width=140 height=54 rowspan=2><img height=54 
      src="images2006/bottom_r.gif" width=140 usemap=#Map 
  border=0></td>
    </tr>
    <tr>
      <td class=sxpta-font2 align=middle height=30><table class=black cellspacing=0 cellpadding=0 width=610 align=center 
      border=0>
          <tbody>
            <tr>
              <td width=170></td>
              <td valign=bottom width=394 height=28></td>
            </tr>
          </tbody>
        </table></td>
    </tr>
  </tbody>
</table>

<%end sub%>
<%
sub addpost()
'    if isre then 
id=request("news_id")
    dwt.out "<div align=center> <hr width=100% size=1>"
	dwt.out "<form method='post' action='news_csyy_view.asp' name='form1' align=center>"
	dwt.out "   <textarea name='body' cols='40' rows='10'></textarea><br/>"	
	dwt.out"<input name='action' type='hidden' value='savepost'> <input name='news_id' type='hidden' value="&id&">      <input  type='submit' name='Submit' value=' �ظ� ' style='cursor:hand;'>"& vbCrLf
    dwt.out "</form></div>	"
'end if   
 
  end sub 
  
  sub del()
dim rsdel2,sqldel2
ID=request("ID3")



	set rsdel2=server.createobject("adodb.recordset")
	sqldel="delete * from csyy_re where id="&id
	rsdel2.open sqldel,conncsyy,1,3
	dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
	'rsdel.close
	set rsdel2=nothing  
end sub

%>
</body>
</html>


