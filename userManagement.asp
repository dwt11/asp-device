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
<!--#include file="inc/md5.asp"-->


<%
dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh
dim rsadd,sqladd,TrueIP,id,rsedit,sqledit,rsdel,sqldel
dim sqluser,rsuser
url="usermanagement.asp"
keys=trim(request("keyword")) 
'groupid=trim(request("group")) 

Dwt.out "<html>"
Dwt.out "<head>"
Dwt.out "<title>�û�����</title>"
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "</head>"
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"
action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
    'message session("userid")&"yyyyy"&session("pagelevelid")&"sdfdsfdf"&truepagelevelh(session("groupid"),2,session("pagelevelid"))
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	
'if request("action1")="editgrouplevel" then call editgrouplevel
sub add()
  Dwt.out"<script type=""text/javascript"" src=""js/regedit.js""></script>"&vbcrlf
 '�����û�
   Dwt.out"<form method='post' action='usermanagement.asp' name='form1' >"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�� �� �� ��</strong></Div></td>    </tr>"
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 Dwt.out"<strong>�� �� ����</strong></td>"
	 Dwt.out"<td width='88%' class='tdbg'>"
	 Dwt.out "<input name='username' type='text' id='input1' onblur='return myuser()' />"
	 Dwt.out "<span id='sps1'></span> "
	 Dwt.out "</td>    </tr>   "
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 
	 
	 
	 Dwt.out"<strong>��&nbsp;&nbsp;&nbsp;&nbsp;����</strong></td>"
	 Dwt.out"<td width='88%' class='tdbg'><input name='username1' type='text'></td>    </tr>   "
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�룺</strong></td> "
	 Dwt.out"<td width='88%' class='tdbg' >"
	 Dwt.out"<input type='password' name='password' id='input2' onblur='return checkpassword()'><span id='sps2'></span></td>    </tr> "
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ȷ�����룺</strong></td> "
	 Dwt.out"<td width='88%' class='tdbg'>"
	 Dwt.out"<input type='password' name='password1'  id='input3' onblur='return checkreturnpass()'/><span id='sps3'></span></td>    </tr> "
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ�� </strong></td>"      
    Dwt.out"<td width='88%' class='tdbg'>"
	Dwt.out"<select name='levelclass' size='1' id='input5' onblur='return checklevelclass()' onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
    Dwt.out "<select name='levelzclass' size='1' >" & vbCrLf
    Dwt.out "<option  selected>ѡ��������</option>" & vbCrLf
    Dwt.out "</select><span id='sps5'></span></td></tr>  "  & vbCrLf
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form1.levelclass.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""���ӷ���"",""0"");" & vbCrLf
		else
		   Dwt.out"group["&rscj("levelid")&"][0]=new Option(""����"",""0"");" & vbCrLf
		do while not rsbz.eof
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




    Dwt.out "var temp=document.form1.levelzclass" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//--></script>" & vbCrLf
		   '����Ȩ����20080330�޸�
   	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����Ȩ���飺 </strong></td>"      
       Dwt.out"<td width='88%' class='tdbg'>"
	   if session("levelclass")=10 then 
			Dwt.out"<select name='groupid' size='1'>"& vbCrLf
			Dwt.out"<option  selected>ѡ������Ȩ����</option>"& vbCrLf
			sqlcj="SELECT * from grouplevel "& vbCrLf
			set rscj=server.createobject("adodb.recordset")
			rscj.open sqlcj,conn,1,1
			do while not rscj.eof
				Dwt.out"<option value='"&rscj("id")&"'"
				'if rscj("id")=rsedit("groupid")then Dwt.out "selected"
				Dwt.out ">"&rscj("name")&"</option>"& vbCrLf
			
				rscj.movenext
			loop
			rscj.close
			set rscj=nothing
			Dwt.out"</select>"  	 & vbCrLf

	   else   
		  Dwt.out "<input  value="&sscjh(rsedit("levelid"))&" type='text' disabled='disabled' >"
	   end if 
	   Dwt.out "</td></tr>"

	
	
	
	    Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='submit' id='submit' value=' �� �� ' style='cursor:hand;' disabled />&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='usermanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from userid" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
	  	  rsadd("username")=ReplaceBadChar(Trim(lcase(Request("username"))))
	  	  rsadd("username1")=ReplaceBadChar(Trim(Request("username1")))
	  rsadd("password")=md5(request("password"),16)
      rsadd("levelid")=ReplaceBadChar(Trim(request("levelclass")))
      rsadd("levelzclass")=ReplaceBadChar(Trim(request("levelzclass")))
      rsadd("groupid")=ReplaceBadChar(Trim(request("groupid")))
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub

sub main()
	sqlbody="SELECT * from userid  "
	'sqlbody="SELECT * from body"
	if keys<>"" then 
		sqlbody=sqlbody&" where username like '%" &keys& "%' "
		title="-���� "&keys
	end if 
	if groupid<>"" then
		sqlbody=sqlbody&" where levelid="&groupid
		title="-"&sscjh(groupid)
	end if 
	sqlbody=sqlbody&" order by levelid aSC,levelzclass asc"

	
	
	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�û�����</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
    	search()
 
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,conn,1,1
      if rsbody.eof and rsbody.bof then 
           Dwt.out "<p align=""center"">��������</p>" 
      else
		  	 Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		  Dwt.out "<form name='form1' method='post' action='usermanagement.asp'>"
		  Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		 Dwt.out "<tr class=""x-grid-header"">"
		 'Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'></Div></td>"
		  Dwt.out "     <td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�û���</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�û�Ȩ��</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>����¼ʱ��</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>����¼IP</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>��¼����</Div></td>"
		  Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>����</Div></td>"
		  Dwt.out "    </tr>"
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
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 'Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center""><input type='checkbox' name='checkuser' value='"&Rsbody("id")&"'></Div></td>"
                 Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh_id&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("username")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("username1")&"</Div></td>"
                  sqllevel="SELECT * from levelname where levelid="&rsbody("levelid")
                 set rslevel=server.createobject("adodb.recordset")
                 rslevel.open sqllevel,conn,1,1
                 if rslevel.eof and rslevel.bof then 
                     Dwt.out "   <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">��������</Div></td>" 
                 else 
                     Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rslevel("levelname")&" "
                 end if
                 rslevel.close
                 set rslevel=nothing
               
			   
			      sqllevel="SELECT * from bzname where id="&rsbody("levelzclass")
                 set rslevel=server.createobject("adodb.recordset")
                 rslevel.open sqllevel,conn,1,1
                 if rslevel.eof and rslevel.bof then 
                     'Dwt.out "   <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">����Ȩ��</Div></td>" 
                 else 
                     Dwt.out rslevel("bzname")&"</Div></td>"
                 end if
                 rslevel.close
                 set rslevel=nothing
                 Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("dldate")&"</Div></td>"
                 Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("dlip")&"</Div></td>"
                 Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsbody("dlcs")&"</Div></td>"
                  Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=""center"">"
				 ' Dwt.out "<a href='usermanagement.asp?action=editpagelevel&checkuser="&rsbody("id")&"'>ҳ��Ȩ������</a>&nbsp;"
				  'Dwt.out "<a href='usermanagement.asp?action=editgrouplevel&checkuser="&rsbody("id")&"'>��Ȩ������</a>&nbsp;"
				  Dwt.out "<a href='usermanagement.asp?action=edit&ID="&rsbody("id")&"'>�༭</a>&nbsp;"
				 if rsbody("levelid")>0 then Dwt.out "  <a href='usermanagement.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('ȷ��Ҫɾ�����û���');"">ɾ��</a></Div></td>"
                 Dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
       end if
       rsbody.close
       set rsbody=nothing
        conn.close
        set conn=nothing
        Dwt.out "</table>"
      	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
        'Dwt.out "<input name='button' type=button onClick='for(i=0;i<=checkuser.length-1;i++){checkuser(i).checked=true}' value='ȫѡ'>"
        'Dwt.out "<input name='button' type=button onClick='for(i=0;i<=checkuser.length-1;i++){checkuser(i).checked=false}' value='ȫ��'>"
        'Dwt.out "<input name='submit' type='submit' value='��������ҳ��Ȩ��' >"
		'Dwt.out "<input type='hidden' name='action' value='editpagelevel'>"
'        Dwt.out "<input name='submit' type='submit' value='����������Ȩ��' >"   ����ͬʱ��������ACTION����Ȩ�����õ��٣���ʱ�Ȳ��������
'		Dwt.out "<input type='hidden' name='action1' value='editgrouplevel'>"
		Dwt.out "</Div></Div>"
		Dwt.out "</form>"
	   call showpage1(page,url,total,record,PgSz)
	   Dwt.out "</Div></Div>"
	  ' else
      '    Dwt.out "<Script Language=Javascript>window.alert('��Ȩ�鿴��ҳ����');history.back()<Script>"
	  ' end if 
end sub




sub edit()
     '�༭�û�
   id=request("id")
   if id="" then id=session("userid")
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from userid where id="&id
   rsedit.open sqledit,conn,1,1

   Dwt.out"<form method='post' action='usermanagement.asp' name='form1'>"
   Dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>�� �� �� ��</strong></Div></td>    </tr>"
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 Dwt.out"<strong>�� �� ����</strong></td>"
   if session("levelCLASS")=0 then 
	'if rsedit("level")=0 then 
		'  Dwt.out"<td width='88%' class='tdbg'><input name='username' type='text' disabled='true'  value='"&rsedit("username")&"'></td>    </tr>   "
 'else
	  Dwt.out"<td width='88%' class='tdbg'><input name='username' type='text' value='"&rsedit("username")&"'></td>    </tr>   "

	' end if 
 else 
	Dwt.out"<td width='88%' class='tdbg'><input name='username' type='text' disabled='disabled' value='"&rsedit("username")&"'><input name='username' type='hidden' value="&rsedit("username")&"></td>    </tr>   "

end if 
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 Dwt.out"<strong>��&nbsp;&nbsp;&nbsp;&nbsp;����</strong></td>"
	  Dwt.out"<td width='88%' class='tdbg'><input name='username1' type='text' value='"&rsedit("username1")&"'></td>    </tr>   "

	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��&nbsp;&nbsp;&nbsp;&nbsp;�룺</strong></td> "
	 Dwt.out"<td width='88%' class='tdbg'>"
	 Dwt.out"<input type='password' name='password' ></td>    </tr> "
	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>ȷ�����룺</strong></td> "
	 Dwt.out"<td width='88%' class='tdbg'>"
	 Dwt.out"<input type='password' name='password1' ></td>    </tr> "
	   
	   
	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>������λ�� </strong></td>"      
       Dwt.out"<td width='88%' class='tdbg'>"
	   if session("levelclass")=10 then 
			Dwt.out"<select name='levelclass' size='1'>"& vbCrLf
			Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
			sqlcj="SELECT * from levelname "& vbCrLf
			set rscj=server.createobject("adodb.recordset")
			rscj.open sqlcj,conn,1,1
			do while not rscj.eof
				Dwt.out"<option value='"&rscj("levelid")&"'"
				if rscj("levelid")=rsedit("levelid")then Dwt.out "selected"
				Dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
			
				rscj.movenext
			loop
			rscj.close
			set rscj=nothing
			Dwt.out"</select>"  	 & vbCrLf

	   else   
		  Dwt.out "<input  value="&sscjh(rsedit("levelid"))&" type='text' disabled='disabled' >"
	   end if 
	  'Dwt.out "</td></td>"
	   'message session("levelclass")
	
	   'Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>�ӷ������ã� </strong></td>"      
       'Dwt.out"<td width='88%' class='tdbg'>"
	   if session("levelclass")=10 then 
			Dwt.out"<select name='levelzclass' size='1'>"& vbCrLf
			Dwt.out"<option value=0 selected>ѡ����������</option>"& vbCrLf
			sqlcj="SELECT * from bzname "& vbCrLf
			set rscj=server.createobject("adodb.recordset")
			rscj.open sqlcj,conn,1,1
			do while not rscj.eof
				Dwt.out"<option value='"&rscj("id")&"'"
				if rscj("id")=rsedit("levelzclass")then Dwt.out "selected"
				Dwt.out ">"&rscj("bzname")&"</option>"& vbCrLf
			
				rscj.movenext
			loop
			rscj.close
			set rscj=nothing
			Dwt.out"</select>"  	 & vbCrLf

	   else   
		  Dwt.out "<input  value='"&ssbzh(rsedit("levelzclass"))&"' type='text' disabled='disabled' >"
	   end if 
	   Dwt.out "</td></tr>"
	   

	   '����Ȩ����20080330�޸�
   	   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����Ȩ���飺 </strong></td>"      
       Dwt.out"<td width='88%' class='tdbg'>"
	   if session("levelclass")=10 then 
			Dwt.out"<select name='groupid' size='1'>"& vbCrLf
			Dwt.out"<option  selected>ѡ������Ȩ����</option>"& vbCrLf
			sqlcj="SELECT * from grouplevel "& vbCrLf
			set rscj=server.createobject("adodb.recordset")
			rscj.open sqlcj,conn,1,1
			do while not rscj.eof
				Dwt.out"<option value='"&rscj("id")&"'"
				if rscj("id")=rsedit("groupid")then Dwt.out "selected"
				Dwt.out ">"&rscj("name")&"</option>"& vbCrLf
			
				rscj.movenext
			loop
			rscj.close
			set rscj=nothing
			Dwt.out"</select>"  	 & vbCrLf

	   else   
		  Dwt.out "<input  value="&sscjh(rsedit("levelid"))&" type='text' disabled='disabled' >"
	   end if 
	   Dwt.out "</td></tr>"



    Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from userid where ID="&ReplaceBadChar(Trim(request("ID")))
	
	rsedit.open sqledit,conn,1,3
	rsedit("username")=Request("username")
	rsedit("username1")=Request("username1")
	if request("password")<>"" then rsedit("password")=md5(request("password"),16)
	if session("levelclass")=10 then 
	 rsedit("groupid")=ReplaceBadChar(Trim(request("groupid")))
	 rsedit("levelid")=ReplaceBadChar(Trim(request("levelclass")))
	 rsedit("levelzclass")=ReplaceBadChar(Trim(request("levelzclass")))
	end if 
	rsedit.update
	rsedit.close
	Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from userid where id="&id
rsdel.open sqldel,conn,1,3
Dwt.out"<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub

Dwt.out "</body></html>"
sub search()
	dim sqlcj,rscj
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<form method='Get' name='SearchForm' action='usermanagement.asp'>" & vbCrLf
	Dwt.out "<a href=""usermanagement.asp?action=add"">����û�</a>&nbsp;&nbsp;�û���������" & vbCrLf
	Dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
	Dwt.out "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
	Dwt.out "�����û��鿴��"
	Dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>������ת����</option>" & vbCrLf
	sqlcj="SELECT * from levelname"& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			Dwt.out"<option value='usermanagement.asp?group="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then Dwt.out" selected"
			Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		Dwt.out "</select>" & vbCrLf
	Dwt.out "</form></Div></Div>" & vbCrLf
end sub

Call CloseConn
%>