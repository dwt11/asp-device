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


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh,xh_id
dim rsadd,sqladd,TrueIP,id,rsedit,sqledit,rsdel,sqldel
dim sqluser,rsuser,sqlcj,rscj
url="grouplevelmanagement.asp"
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
	  case "editpagelevel"
    call editpagelevel
  case "saveeditpagelevel"
    call saveeditpagelevel
  case "editgrouplevel"
    call editgrouplevel
  case "saveeditgrouplevel"
    call saveeditgrouplevel

end select	
dwt.out "<html>"
dwt.out "<head>"
dwt.out "<title>Ȩ�������</title>"
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function CheckAdd(){" & vbCrLf
 dwt.out " if(document.form1.bzname.value==''){" & vbCrLf
dwt.out "      alert('��������Ϊ�գ�');" & vbCrLf
dwt.out "   document.form1.bzname.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

'dwt.out "  if(document.form1.sscj.value==''){" & vbCrLf
'dwt.out "      alert('δѡ���������䣡');" & vbCrLf
' dwt.out "  document.form1.sscj.focus();" & vbCrLf
'dwt.out "      return false;" & vbCrLf
'dwt.out "    }" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"

sub add()
   '�����û�
   dwt.out"<form method='post' action='grouplevelmanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�� �� Ȩ �� ��</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>Ȩ�������ƣ�</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'><input name='grouplevelname' type='text'></td>    </tr>   "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out"<strong>������</strong></td>"
	 dwt.out"<td width='88%' class='tdbg'>  <textarea name='grouplevelinfo' cols='20' rows='10'></textarea></td>    </tr>   "
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='grouplevelmanagement.asp';"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from grouplevel" 
      rsadd.open sqladd,conn,1,3
      rsadd.addnew
      rsadd("name")=ReplaceBadChar(Trim(Request("grouplevelname")))
      rsadd("info")=ReplaceBadChar(Trim(request("grouplevelinfo")))
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub

sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>Ȩ�������</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
    dwt.out "<div class='x-toolbar'>" & vbCrLf
    dwt.out "<div align=left><a href=grouplevelmanagement.asp?action=add>���Ȩ����</a></div>"
	dwt.out "</div>"
	  sqlbody="SELECT * from grouplevel"
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,conn,1,1
      if rsbody.eof and rsbody.bof then 
           dwt.out "<p align=""center"">��������</p>" 
      else
		  	 dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
			dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			dwt.out "<tr  class=""x-grid-header"">" 
			dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>Ȩ��������</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
			dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
			dwt.out "    </tr>"
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
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("name")&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rsbody("info")&" &nbsp;</div></td>"

				  dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
                 Dwt.out "<a href='grouplevelmanagement.asp?name="&rsbody("name")&"&action=editpagelevel&groupid="&rsbody("id")&"'>ҳ��Ȩ������</a>&nbsp;"
				  Dwt.out "<a href='grouplevelmanagement.asp?name="&rsbody("name")&"&action=editgrouplevel&groupid="&rsbody("id")&"'>��Ȩ������</a>&nbsp;"
				  dwt.out "<a href='grouplevelmanagement.asp?action=edit&ID="&rsbody("id")&"'>�༭</a>&nbsp;"
				 dwt.out "  <a href='grouplevelmanagement.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('ȷ��Ҫɾ��������');"">ɾ��</a></div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
       end if
       rsbody.close
       set rsbody=nothing
        conn.close
        set conn=nothing
        dwt.out "</table>"
       call showpage1(page,url,total,record,PgSz)
	   dwt.out "</div></div>"
end sub



sub editgrouplevel()
    'checkvalue=request.form("checkuser")
    'if checkvalue="" then 
	'dim checkvalue,leftmdb
	dim groupid,leftmdb,sql,rs,groupname
	groupid=request("groupid")
	groupname=request("name")
    'checkvalue1=request.form("checkuser")'���ڱ�������ʱ����saveeditlevel�����û���
	'if checkvalue1="" then checkvalue1=request("checkuser")'���ڱ�������ʱ����saveeditlevel�����û���
	'if checkvalue="" then 
    '     if checkvalue="" then Dwt.out"<Script Language=Javascript>window.alert('������ѡ��һ������');history.back()"
    'else
         'checkvalue=split(checkvalue,",")
	     'For i = LBound(checkvalue) To UBound(checkvalue)
	    '    username=username&usernameh(checkvalue(i))&"&nbsp;&nbsp;"
		' Next 
		Dwt.out "<Div class='pre'><Div align='center'>" & vbCrLf
		Dwt.out "���� <font color='#ff0000'>"&groupname&"</font> ��ĵ�λȨ��"
		Dwt.out "</Div></Div>"
		
'		leftmdb="ybdata/left.mdb"
'		Set connleft = Server.CreateObject("ADODB.Connection")
'		connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
'		connleft.Open connl    
		
		  sql="SELECT * from levelname where istq=false"
		  set rs=server.createobject("adodb.recordset")
		  rs.open sql,conn,1,1
		  if rs.eof and rs.bof then 
			 message "���κε�λ" 
		  else
			  Dwt.out "<form name='form1' method='post' action='grouplevelmanagement.asp'>"
			 Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
			 Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			 Dwt.out "<tr class=""x-grid-header"">"
			 Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>������</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�༭</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>ɾ��</Div></td>"
			 Dwt.out "    </tr>"
		  
		  do while not rs.eof 
					dim xh,xh_id
					xh=xh+1
					if xh mod 2 =1 then 
					  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rs("levelname")&"&nbsp;</Div></td>"
					
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
					Dwt.out "<input type='checkbox' name='check_edit' value='"&rs("levelid")&"'"
					call checkgrouplevelh(groupid,0,rs("levelid"))
					Dwt.out ">"
					 '   response.Write(groupid&"sdfsdfsdf"&rs("levelid"))

					Dwt.out "</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"& vbCrLf
					Dwt.out "<input type='checkbox' name='check_del' value='"&rs("levelid")&"'"
					call checkgrouplevelh(groupid,1,rs("levelid"))
					Dwt.out ">"& vbCrLf
					Dwt.out "</Div></td>"& vbCrLf
					
					Dwt.out "</tr>"
				
			rs.movenext
			loop
			 Dwt.out "</table>"
			  Dwt.out "<Div class='x-toolbar'>" & vbCrLf
			  Dwt.out"			  <input name='action' type='hidden' value='saveeditgrouplevel'> <input name='groupid' type='hidden' value='"&groupid&"'>    <input  type='submit' name='Submit' value='�� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
			  Dwt.out "</Div>"
			 Dwt.out "</Div>"
		end if 
		  rs.close
		  set rs=nothing
		  Dwt.out "</Div>"
		  Dwt.out "</form>"
	'end if
end sub
sub saveeditgrouplevel()
	dim groupid,checkuser,rsedit,sqledit
	groupid=request("groupid")
	checkuser=split(checkuser,",")
	'For i = LBound(checkuser) To UBound(checkuser)
		set rsedit=server.createobject("adodb.recordset")
		sqledit="select * from grouplevel where ID="&groupid
		rsedit.open sqledit,conn,1,3
        'message Request("check_display")&"/"&Request("check_new")&"/"&Request("check_edit")&"/"&Request("check_del")
		rsedit("grouplevel")=Request("check_edit")&"/"&Request("check_del")
		rsedit.update
		rsedit.close
	'Next 
	
	Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"


end sub

sub editpagelevel()
    'checkvalue=request.form("checkuser")
    'if checkvalue="" then checkvalue=request("checkuser")
    'checkvalue1=request.form("checkuser")'���ڱ�������ʱ����saveeditlevel�����û���
	'if checkvalue1="" then checkvalue1=request("checkuser")'���ڱ�������ʱ����saveeditlevel�����û���
	'if checkvalue="" then 
     '    if checkvalue="" then Dwt.out"<Script Language=Javascript>window.alert('������ѡ��һ������');history.back()"
    'else
         'checkvalue=split(checkvalue,",")
	     'For i = LBound(checkvalue) To UBound(checkvalue)
	        'username=username&usernameh(checkvalue(i))&"&nbsp;&nbsp;"
		' Next 
	dim groupid,leftmdb,sql,rs,groupname,connleft,connl
	groupid=request("groupid")
	groupname=request("name")
		Dwt.out "<Div class='pre'><Div align='center'>" & vbCrLf
		Dwt.out "���������û���ҳ��Ȩ�ޣ�"&groupname&"<br/><font color='#ff0000'>����ϵͳ���������ü̳�Ȩ�ޣ���ʹһ���������������ӷ������ͬ��������ͬ��Ȩ��</font><br/><font color='#ff0000'>�޸��û�Ȩ�޺��û������µ�¼��������Ч</font>"
		Dwt.out "</Div></Div>"
		
		leftmdb="ybdata/left.mdb"
		Set connleft = Server.CreateObject("ADODB.Connection")
		connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
		connleft.Open connl    
		
		  sql="SELECT * from left_class where zclass=0 order by orderby aSC"
		  set rs=server.createobject("adodb.recordset")
		  rs.open sql,connleft,1,1
		  if rs.eof and rs.bof then 
			 message "���κ���Ŀ" 
		  else
			  Dwt.out "<form name='form1' method='post' action='grouplevelmanagement.asp'>"
			 Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
			 Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			 Dwt.out "<tr class=""x-grid-header"">"
			 Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'>���</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>��Ŀ����</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>������Ŀ</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�鿴</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�½�</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>�༭</Div></td>"
			 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>ɾ��</Div></td>"
			 Dwt.out "    </tr>"
		  
		  do while not rs.eof 
					dim xh,xh_id
					xh=xh+1
					if xh mod 2 =1 then 
					  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					else
					  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
					end if 
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh&"</Div></td>"
					Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rs("name")&"&nbsp;</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">һ��</Div></td>"
					
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
					Dwt.out "<input type='checkbox' name='check_display' value='"&rs("id")&"' "
					    call checkpagelevelh(groupid,0,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					    'call checkpagelevelh(31,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					Dwt.out "/>"
					Dwt.out "</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
					Dwt.out "<input type='checkbox' name='check_new' value='"&rs("id")&"' "
					    call checkpagelevelh(groupid,1,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					    'call checkpagelevelh(31,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					Dwt.out "/>"
					Dwt.out "</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
					Dwt.out "<input type='checkbox' name='check_edit' value='"&rs("id")&"' "
					    call checkpagelevelh(groupid,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					    'call checkpagelevelh(31,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					Dwt.out "/>"
					Dwt.out "</Div></td>"
					Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
					Dwt.out "<input type='checkbox' name='check_del' value='"&rs("id")&"' "
					    call checkpagelevelh(groupid,3,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					    'call checkpagelevelh(31,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
					Dwt.out "/>"
					Dwt.out "</Div></td>"
					
					Dwt.out "</tr>"
							'����
					dim sqlz,rsz
					sqlz="SELECT * from left_class where zclass="&rs("id")&" order by orderby aSC"& vbCrLf
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,connleft,1,1
					if rsz.eof and rsz.bof then 
					else
							dim xhz
						xhz=0
						do while not rsz.eof
							'xh=xh+1
							xhz=xhz+1
							if xh mod 2 =1 then 
							  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
							else
							  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
							end if 
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&xh&"-"&xhz&"</Div></td>"
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsz("name")&"&nbsp;</Div></td>"
							Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rs("name")&"-����</Div></td>"
							
							if connleft.Execute("SELECT isbiglevel FROM left_class WHERE id="&rs("id"))(0) then 
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_display'  disabled='disabled'"
								call checkpagelevelh(groupid,0,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_new'  disabled='disabled'"
								call checkpagelevelh(groupid,1,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_edit' disabled='disabled'"
								call checkpagelevelh(groupid,2,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_del'  disabled='disabled'"
								call checkpagelevelh(groupid,3,rs("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
							else
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_display' value='"&rsz("id")&"'"
								call checkpagelevelh(groupid,0,rsz("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_new' value='"&rsz("id")&"'"
								call checkpagelevelh(groupid,1,rsz("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_edit' value='"&rsz("id")&"'"
								call checkpagelevelh(groupid,2,rsz("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
								Dwt.out " <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
								Dwt.out "<input type='checkbox' name='check_del' value='"&rsz("id")&"'"
								call checkpagelevelh(groupid,3,rsz("id")) '�ж��û���Щҳ����ʾȨ�ޣ��������ʾ�����'��ѡ��'
								Dwt.out "/>"
								Dwt.out "</Div></td>"
							end if 	
						   
						   
						   
						   
						   
						   
						    Dwt.out "</tr>"
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing
				
			rs.movenext
			loop
			 Dwt.out "</table>"
			  Dwt.out "<Div class='x-toolbar'>" & vbCrLf
			  Dwt.out"			  <input name='action' type='hidden' value='saveeditpagelevel'> <input name='groupid' type='hidden' value='"&groupid&"'>    <input  type='submit' name='Submit' value='�� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
			  Dwt.out "</Div>"
			 Dwt.out "</Div>"
		end if 
		  rs.close
		  set rs=nothing
		  Dwt.out "</Div>"
		  Dwt.out "</form>"
	'end if
end sub

sub saveeditpagelevel()
	'checkuser=request("checkuser1")
	'checkuser=split(checkuser,",")
	dim groupid
	groupid=request("groupid")
	'For i = LBound(checkuser) To UBound(checkuser)
		set rsedit=server.createobject("adodb.recordset")
		sqledit="select * from grouplevel where ID="&groupid
		rsedit.open sqledit,conn,1,3
        'message Request("check_display")&"/"&Request("check_new")&"/"&Request("check_edit")&"/"&Request("check_del")
		rsedit("pagelevel")=Request("check_display")&"/"&Request("check_new")&"/"&Request("check_edit")&"/"&Request("check_del")
		rsedit.update
		rsedit.close
	'Next 
	
	Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"


end sub

sub edit()
     '�༭�û�
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from bzname where id="&id
   rsedit.open sqledit,conn,1,1

   dwt.out"<form method='post' action='grouplevelmanagement.asp' name='form1' onsubmit='javascript:return CheckAdd();'>"
   dwt.out"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>�� �� �� ��</strong></div></td>    </tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>��������</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'>"
	 dwt.out"<input type='text' name='bzname' value="&rsedit("bzname")&"></td>    </tr> "
	dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"& vbCrLf      
    dwt.out"<td width='88%' class='tdbg'>"& vbCrLf
	
	dwt.out"<select name='sscj' size='1'>"& vbCrLf
    dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'"
		if rsedit("sscj")=rscj("levelid") then dwt.out"selected"
		dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>"  	 & vbCrLf
    dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>	<input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""history.back();"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from bzname where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conn,1,3
rsedit("bzname")=ReplaceBadChar(Trim(Request("bzname")))
rsedit("sscj")=ReplaceBadChar(Trim(Request("sscj")))
rsedit.update
rsedit.close
	
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from bzname where id="&id
rsdel.open sqldel,conn,1,3
dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub

dwt.out "</body></html>"

Call CloseConn
%>