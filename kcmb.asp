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
dim lxclassid,url,record,pgsz,total,page,start,rowcount,ii,pagename
dim sqlbody,rsbody,xh
dim rs,sql
lxclassid = Trim(Request("lxclassid"))
if lxclassid="" then lxclassid=1
url="ylb.asp?lxclassid="&lxclassid
'��ȡ���࣬�����ڱ���
sql="SELECT * from lxclass where lxznum=0 and lxnum="&lxclassid& vbCrLf
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	pagename=rs("lxname")
rs.close
set rs=nothing

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title> ������������ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function CheckSearch(){" & vbCrLf
response.write "  if(document.SearchForm.lxclassid.value==''){" & vbCrLf
response.write "      alert('�豸���಻��Ϊ�գ�');" & vbCrLf
response.write "  document.SearchForm.lxclassid.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>"&pagename&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'>"
sql="SELECT * from lxclass where lxznum=0"& vbCrLf
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
	response.write "<a href=ylb.asp?lxclassid="&rs("lxnum")&">"&rs("lxname")&"</a>&nbsp;|&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit'  onclick=""window.location.href='/ylb_ned.asp?action=add&lxclassid="&lxclassid&"'""value='���"&pagename&"'>"
response.write "</td>"& vbCrLf
response.write "  </tr>"& vbCrLf
response.write "</table>"& vbCrLf

call search(lxclassid)
if lxclassid<>"" then 
   select case lxclassid
     case 1
		Call djdylb
     case 2
		Call bsq
	 case 3
        Call zhq
     case 4
        Call tjffj  
     case 5
	    Call dcf
	case 6
	    Call djdylb
    case 7
	    Call djdylb		
	case 8
	    Call llycyj		
	case 9
	    Call cwycyj		
	case 10
	    Call jztt				
	case 11
	    Call fxyb	
	case 12
	    Call kt		
	case 13
	    Call pdc		
    case 14
	    Call tjf	
	case 15
	    Call ddzxjg	
   end select
end if 


Sub djdylb()
  sqlbody="SELECT * from ylbbody where dclass="&lxclassid&" order by  id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">��������</p>" 
  else
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
  
     response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
  response.write "<tr class=""title"">"
  response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center""><strong>���</strong></div></td>"
  response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>����</strong></div></td>"
  response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><div align=""center""><strong>λ ��</strong></div></td>"
  response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ͺ�</strong></div></td>"
  response.write "      <td width=""14%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������Χ</strong></div></td>"
  response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>"
  response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ��</strong></div></td>"
  response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ע</strong></div></td>"
  response.write "      <td width=""22%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ ��</strong></div></td>"
  response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        xh=xh+1
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center"">"&xh&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&sscjh_d(rsbody("sscj"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" width=""15%"">"&rsbody("wh")&"</td>"
        response.write "  <td width=""15%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("ggxh")&"&nbsp;</td>"
        response.write "  <td width=""14%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("clfw")&"&nbsp;</td>"
        response.write "  <td width=""11%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("changj")&"&nbsp;</td>"
        response.write " <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("shul")&"&nbsp;</div></td>"
        response.write " <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("whbeizhu")&"&nbsp;</div></td>"
       response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	  call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��
       response.write "</div></td></tr>"
        RowCount=RowCount-1
    rsbody.movenext
    loop
   response.write "</table>"
  call showpage(page,url,total,record,PgSz)
 end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub

Sub bsq()
    dim bsqname,zclass
   response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
   response.write "<tr class='title'>"
   response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center""><strong>���      </strong></div></td>"
   response.write "      <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>����</strong></div></td>"
   response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""9%""><div align=""center""><strong>λ ��</strong></div></td>"
   response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ��</strong></div></td>"
   response.write "      <td width=""24%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ͺ�</strong></div></td>"
   response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
   response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�¶�</strong></div></td>"
   response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѹ��</strong></div></td>"
   response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��/������</strong></div></td>"
   response.write "      <td width=""9%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������Χ</strong></div></td>"
   response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
   response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ע</strong></div></td>"
   response.write "      <td width=""15%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ ��</strong></div></td>"
   response.write "    </tr>"
   sqlbody="SELECT * from ylbbody where dclass=2  order by  id DESC"
   set rsbody=server.createobject("adodb.recordset")
   rsbody.open sqlbody,conn,1,1
   if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
   else
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
       dim start
       start=PgSz*Page-PgSz+1
       rowCount = rsbody.PageSize
       do while not rsbody.eof and rowcount>0
            select case rsbody("zclass")
               case 1
                  bsqname="ѹ��������"
               case 2
                   bsqname="Һλ������"
               case 3
                   bsqname="����������"
               case 4
                   bsqname="ѹ�������"
               case 5
                    bsqname="����˫������׶����"
               case 6
                     bsqname="��λ������"
               case 7
                     bsqname="������������"
               case 8
                     bsqname="������Һλ��"
               case 9
                     bsqname="��Ͳ"
                case 10
                     bsqname="���������"
               case 11
                    bsqname="�¶ȱ�����"
               case 12
                    bsqname="�״�Һλ��"
               case 13
                    bsqname="�н�������"
               case 14
                    bsqname="����������"
               case 15
                    bsqname="����Һλ��"
               case 16
                     bsqname="λ���񶯱�����"
			   case 0
			         bsqname="������"
            end select	
         	xh=xh+1
       		response.write "<tr class=""tdbg""  onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
            response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""3%""><div align=""center"">"&xh&"</div></td>"
            response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center"">"&sscjh_d(rsbody("sscj"))&"</div></td>"
            response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""9%"">"&rsbody("wh")&"&nbsp;</td>"
            response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px"">"&bsqname&"&nbsp;"&zclass&"</td>"
            response.write "      <td width=""24%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("ggxh")&"&nbsp;</td>"
            response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("gyjz")&"&nbsp;</td>"
            response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("czwd")&"&nbsp;</td>"
            response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("czyl")&"&nbsp;</td>"
            response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("qdlx")&"&nbsp;</td>"
            response.write "      <td width=""9%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("clfw")&"&nbsp;</td>"
            response.write "      <td width=""4%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("changj")&"&nbsp;</td>"
            response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("whbeizhu")&"&nbsp;</div></td>"
            response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��
            response.write "</div></td></tr>"
            RowCount=RowCount-1
        rsbody.movenext
        loop
   end if
   rsbody.close
   set rsbody=nothing
   conn.close
   set conn=nothing
   response.write "</table>"
   call showpage(page,url,total,record,PgSz)
end sub


sub zhq()
    response.write"<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='1'><tr class='title'>" & vbCrLf
	response.write"<td  style='border-bottom-style: solid;border-width:1px' width='4%'><div align='center'><strong>���      </strong></div></td>" & vbCrLf
    response.write"      <td  style='border-bottom-style: solid;border-width:1px' width='4%'><div align='center'><strong>����</strong></div></td>" & vbCrLf
    response.write"    <td style='border-bottom-style: solid;border-width:1px' width='13%'><div align='center'><strong>λ ��</strong></div></td>" & vbCrLf
    response.write"     <td width='23%' style='border-bottom-style: solid;border-width:1px'><div align='center'><strong>����ͺ�</strong></div></td>" & vbCrLf
    response.write"     <td width='10%' style='border-bottom-style: solid;border-width:1px'><div align='center'><strong>��������</strong></div></td>" & vbCrLf
    response.write"     <td width='14%' style='border-bottom-style: solid;border-width:1px'><div align='center'><strong>�� ��</strong></div></td>" & vbCrLf
    response.write"     <td width='8%' style='border-bottom-style: solid;border-width:1px'><div align='center'><strong>�� ע</strong></div></td>" & vbCrLf
    response.write"     <td width='24%' style='border-bottom-style: solid;border-width:1px'><div align='center'><strong>ѡ ��</strong></div></td>" & vbCrLf
    response.write"   </tr>" & vbCrLf
	sqlbody="SELECT * from ylbbody where dclass=3 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
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
       dim start
       start=PgSz*Page-PgSz+1
       rowCount = rsbody.PageSize
       do while not rsbody.eof and rowcount>0
           xh=xh+1
           response.write"<tr class='tdbg'  onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
           response.write"<td  style='border-bottom-style: solid;border-width:1px' width='4%'><div align='center'>"&xh&"</div></td>" & vbCrLf
           response.write"      <td  style='border-bottom-style: solid;border-width:1px' width='4%'><div align='center'>"&sscjh_d(rsbody("sscj"))&"</div></td>" & vbCrLf
           response.write"      <td style='border-bottom-style: solid;border-width:1px' width='13%'>"&rsbody("wh")&"</td>" & vbCrLf
           response.write"      <td width='23%' style='border-bottom-style: solid;border-width:1px'>"&rsbody("ggxh")&"&nbsp;</td>" & vbCrLf
           response.write"      <td width='10%' style='border-bottom-style: solid;border-width:1px'>"&rsbody("changj")&"&nbsp;</td>" & vbCrLf
           response.write"      <td width='14%' style='border-bottom-style: solid;border-width:1px'>"&rsbody("qdlx")&"&nbsp;</td>" & vbCrLf
           response.write"      <td width='8%' style='border-bottom-style: solid;border-width:1px'><div align='center'>"&rsbody("whbeizhu")&"&nbsp;</div></td>" & vbCrLf
          response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
     	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��
           response.write "</div></td></tr>"
           RowCount=RowCount-1
	  rsbody.movenext
      loop
  end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
  response.write"</table>" & vbCrLf
  call showpage(page,url,total,record,PgSz)
end sub


sub tjffj()
 response.write"<table width=100%  border=0 align=center cellpadding=0 cellspacing=1>" & vbCrLf
 response.write"<tr class=title>" & vbCrLf
 response.write"<td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=center><strong>���</strong></div></td>" & vbCrLf
 response.write"     <td  style='border-bottom-style: solid;border-width:1px' width='4%'><div align=center><strong>����</strong></div></td>" & vbCrLf
 response.write"     <td style='border-bottom-style: solid;border-width:1px' width='13%'><div align=center><strong>λ ��</strong></div></td>" & vbCrLf
 response.write"     <td width='17%' style='border-bottom-style: solid;border-width:1px'><div align=center><strong>����ͺ�</strong></div></td>" & vbCrLf
 response.write"     <td width='16%' style='border-bottom-style: solid;border-width:1px'><div align=center><strong>��������</strong></div></td>" & vbCrLf
 response.write"     <td width='14%' style='border-bottom-style: solid;border-width:1px'><div align=center><strong>�� ��</strong></div></td>" & vbCrLf
 response.write"     <td width='8%' style='border-bottom-style: solid;border-width:1px'><div align=center><strong>�� ע</strong></div></td>" & vbCrLf
 response.write"     <td width='24%' style='border-bottom-style: solid;border-width:1px'><div align=center><strong>ѡ ��</strong></div></td>" & vbCrLf
 response.write"   </tr>"
 sqlbody="SELECT * from ylbbody where dclass=4 order by  id DESC"
 set rsbody=server.createobject("adodb.recordset")
 rsbody.open sqlbody,conn,1,1
 if rsbody.eof and rsbody.bof then 
    response.write "<p align='center'>��������</p>" 
 else
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
    dim start
    start=PgSz*Page-PgSz+1
    rowCount = rsbody.PageSize
    do while not rsbody.eof and rowcount>0
        xh=xh+1
        response.write "<tr  class=""tdbg""  onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
      response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""center"">"&xh&"</div></td>"
      response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""4%""><div align=""cente"">"&sscjh_d(rsbody("sscj"))&"</div></td>"
      response.write "<td style=""border-bottom-style: solid;border-width:1px"" width=""13%"">"&rsbody("wh")&"</td>"
      response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("ggxh")&"&nbsp;</td>"
      response.write "<td width=""16%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("changj")&"&nbsp;</td>"
      response.write "<td width=""14%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("llname")&"&nbsp;</td>"
      response.write "<td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("whbeizhu")&"&nbsp;</div></td>"
      response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	  call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��
      response.write "</div></td></tr>"
      RowCount=RowCount-1
	  rsbody.movenext
      loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
response.write "</table>"
call showpage(page,url,total,record,PgSz)
end sub

sub dcf()
 response.write "<table width=100%  border=1 align=center cellpadding=0 cellspacing=0 bordercolor=#FFFFFF bgcolor=#CCCCCC>" & vbCrLf
 response.write "   <tr class=title>" & vbCrLf
 response.write "     <td   width=4% rowspan=2><div align=center>���      </div></td>" & vbCrLf
 response.write "     <td  width=6% rowspan=2><div align=center>����</div></td>" & vbCrLf
 response.write "     <td width=13% rowspan=2><div align=center>λ ��</div></td>" & vbCrLf
 response.write "     <td colspan=3 ><div align=center>��Ȧ</div></td>" & vbCrLf
 response.write "     <td colspan=3 ><div align=center>����</div></td>" & vbCrLf
 response.write "     <td width=6% rowspan=2 ><div align=center>�� ע</div></td>" & vbCrLf
 response.write "     <td width=17% rowspan=2 ><div align=center>ѡ ��</div></td>" & vbCrLf
 response.write "   </tr>" & vbCrLf
 response.write "    <tr class=title>" & vbCrLf
 response.write "     <td width=12% ><div align=center>�ͺ�</div></td>" & vbCrLf
 response.write "     <td width=4% ><div align=center>����</div></td>" & vbCrLf
 response.write "     <td width=15% ><div align=center>����</div></td>" & vbCrLf
 response.write "     <td width=10% ><div align=center>�ͺ�</div></td>" & vbCrLf
 response.write "     <td width=7% ><div align=center>ͨ·</div></td>" & vbCrLf
 response.write "     <td width=6% ><div align=center>����</div></td>" & vbCrLf
 response.write "   </tr>" & vbCrLf
 sqlbody="SELECT * from ylbbody where dclass=5 order by  id DESC"
 set rsbody=server.createobject("adodb.recordset")
 rsbody.open sqlbody,conn,1,1
 if rsbody.eof and rsbody.bof then 
    response.write "<p align='center'>��������</p>" 
 else
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
	dim start
	start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
     xh=xh+1   
     response.write "<tr  class=""tdbg""  onmouseout=""'this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
      response.write "<td width=""4%""><div align=""center"">"&(xh)&"</div></td>"& vbCrLf
     response.write " <td width=""6%""><div align=""center"">"&sscjh_d(rsbody("sscj"))&"</div></td>"& vbCrLf
     response.write " <td width=""13%"">"&rsbody("wh")&"</td>"& vbCrLf
     response.write " <td width=""12%"">"&rsbody("xianxh")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""4%"">"&rsbody("xiangd")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""15%"">"&rsbody("xiangcj")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""10%"">"&rsbody("fatixh")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""7%"">"&rsbody("fatitl")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""6%"">"&rsbody("faticj")&"&nbsp;</td>"& vbCrLf
     response.write " <td width=""6%""><div align=""center"">"&rsbody("whbeizhu")&"&nbsp;</div></td>"& vbCrLf
     response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	  call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��
      response.write "</div></td></tr>"& vbCrLf
      RowCount=RowCount-1
      rsbody.movenext
   loop
 end if
 rsbody.close
 set rsbody=nothing
 conn.close
 set conn=nothing
 response.write "</table>"
 call showpage(page,url,total,record,PgSz)
end sub

Sub llycyj()
      response.write "<table width=100%  border=1 align=center cellpadding=0 cellspacing=0 bordercolor=#FFFFFF bgcolor=#CCCCCC>"& vbCrLf
    response.write "<tr class=title>"& vbCrLf
    response.write "  <td   width=4% rowspan=2><div align=center>���      </div></td>"& vbCrLf
    response.write "  <td  width=5% rowspan=2><div align=center>����</div></td>"& vbCrLf
    response.write "  <td width=11% rowspan=2><div align=center>λ ��</div></td>"& vbCrLf
    response.write "  <td width=10% rowspan=2><div align=center>����</div></td>"& vbCrLf
    response.write "  <td width=7% rowspan=2><div align=center>ȡѹ��ʽ</div></td>"& vbCrLf
    response.write "  <td width=13% rowspan=2><div align=center>��ѹ��Χ(Kpa)</div></td>"& vbCrLf
    response.write "  <td colspan=3 ><div align=center>һ��Ԫ���ߴ�(mm)</div></td>"& vbCrLf
    response.write "  <td width=10% rowspan=2 ><div align=center>�� ע</div></td>"& vbCrLf
    response.write "  <td width=19% rowspan=2 ><div align=center>ѡ ��</div></td>"& vbCrLf
    response.write "</tr>"& vbCrLf
    response.write "<tr class=title>"& vbCrLf
   response.write "   <td width=7% ><div align=center>�׾�</div></td>"& vbCrLf
   response.write "   <td width=7% ><div align=center>�⾶</div></td>"& vbCrLf
   response.write "   <td width=7% ><div align=center>���</div></td>"& vbCrLf
   response.write "   </tr>"& vbCrLf
   sqlbody="SELECT * from ylbbody where dclass=8 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
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
	dim start
	start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
   xh=xh+1%>
           
		   
		   
    <tr class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'" >
      <td width="4%"><div align="center"><%=(xh)%></div></td>
      <td width="5%"><div align="center"><%=(sscjh_d(rsbody("sscj")))%></div></td>
      <td width="11%"><%=rsbody("wh")%>&nbsp;</td>
      <td width="10%"><%=rsbody("llname")%>&nbsp;</td>
      <td width="7%"><%=rsbody("qyfs")%>&nbsp;</td>
      <td width="13%"><%=rsbody("clfw")%>&nbsp;</td>
      <td width="7%"><%=rsbody("llkj")%>&nbsp;</td>
      <td width="7%"><%=rsbody("llwj")%>&nbsp;</td>
      <td width="7%"><%=rsbody("llhd")%>&nbsp;</td>
      <td width="10%"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
  </table>
<%call showpage(page,url,total,record,PgSz)

end sub

Sub cwycyj()
  sqlbody="SELECT * from ylbbody where dclass=9 order by  id DESC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     response.write "<p align=""center"">��������</p>" 
  else
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
     response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
  response.write "<tr class=""title"">"
  response.write "<td  style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center""><strong>���</strong></div></td>"
  response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><div align=""center""><strong>����</strong></div></td>"
  response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""8%""><div align=""center""><strong>λ ��</strong></div></td>"
  response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ��</strong></div></td>"
  response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ͺ�</strong></div></td>"
  response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����</strong></div></td>"
  response.write "      <td width=""10%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�ߴ�</strong></div></td>"
  response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>������Χ</strong></div></td>"
  response.write "      <td width=""8%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��������</strong></div></td>"

  response.write "      <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ��</strong></div></td>"
  response.write "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�� ע</strong></div></td>"
  response.write "      <td width=""22%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ ��</strong></div></td>"
  response.write "    </tr>"
  
  do while not rsbody.eof and rowcount>0
        xh=xh+1
        response.write " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" width=""6%""><div align=""center"">"&xh&"</div></td>"
        response.write " <td  style=""border-bottom-style: solid;border-width:1px"" width=""8%""><div align=""center"">"&sscjh_d(rsbody("sscj"))&"</div></td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" width=""8%"">"&rsbody("wh")&"</td>"
        response.write "  <td style=""border-bottom-style: solid;border-width:1px"" width=""8%"">"&rsbody("llname")&"</td>"
        response.write "  <td width=""8%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("ggxh")&"&nbsp;</td>"
        response.write "  <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("llcz")&"&nbsp;</td>"
        response.write "  <td width=""10%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("llhd")&"&nbsp;</td>"
        response.write "  <td width=""8%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("clfw")&"&nbsp;</td>"
        response.write "  <td width=""8%"" style=""border-bottom-style: solid;border-width:1px"">"&rsbody("changj")&"&nbsp;</td>"
        response.write " <td width=""5%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("shul")&"&nbsp;</div></td>"
        response.write " <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsbody("whbeizhu")&"&nbsp;</div></td>"
       response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

       response.write "</div></td></tr>"
        RowCount=RowCount-1
    rsbody.movenext
    loop
   response.write "</table>"
  call showpage(page,url,total,record,PgSz)
 end if
  rsbody.close
  set rsbody=nothing
  conn.close
  set conn=nothing
end sub


sub jztt()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><strong>���      </strong></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><strong>����</strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="13%"><div align="center"><strong>λ ��</strong></div></td>
      <td width="17%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>̽ͷ����ͺ�</strong></div></td>
      <td width="16%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>��������ͺŹ��</strong></div></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ӽ����ͺŹ��</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ע</strong></div></td>
      <td width="24%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѡ ��</strong></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=10 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
  
  xh=xh+1%>
           
		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><%=(xh)%></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><%=(sscjh_d(rsbody("sscj")))%></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="13%"><%=rsbody("wh")%></td>
      <td width="17%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ggxh")%>&nbsp;</td>
      <td width="16%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("yanxh")%></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("jiejxh")%>&nbsp;</td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("changj")%></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</table>


<%call showpage(page,url,total,record,PgSz)

end sub

sub fxyb()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><strong>���      </strong></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><strong>����</strong></div></td>
	  <td  style="border-bottom-style: solid;border-width:1px" width="12%"><div align="center"><strong>����</strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="12%"><div align="center"><strong>λ ��</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>������Χ</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>��������</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�������ͺ�</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�缫̽ͷ�ͺ�</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="3%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ע</strong></div></td>
      <td width="15%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѡ ��</strong></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=11 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
  xh=xh+1%>
           
		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      
	  <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><%=(xh)%></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="4%"><div align="center"><%=(sscjh_d(rsbody("sscj")))%></div></td>
	  <td  style="border-bottom-style: solid;border-width:1px" width="12%"><%=rsbody("llname")%></td>
      <td style="border-bottom-style: solid;border-width:1px" width="12%"><%=rsbody("wh")%></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("clfw")%>&nbsp;</td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("gyjz")%></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("fenbsq")%></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("fendj")%>&nbsp;</td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("changj")%></td>
      <td width="3%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</table>


<%call showpage(page,url,total,record,PgSz)
end sub

sub kt()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td width="5%" height="20"  style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���      </strong></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="16%"><div align="center"><strong>����</strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="15%"><div align="center"><strong>����ͺ�</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѹ��������</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���䷽ʽ</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�������</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ע</strong></div></td>
      <td width="16%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѡ ��</strong></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=12 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
  
	xh=xh+1%>
           
		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td  style="border-bottom-style: solid;border-width:1px" width="5%"><div align="center"><%=(xh)%></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="16%"><%=rsbody("llname")%></td>
      <td style="border-bottom-style: solid;border-width:1px" width="15%"><%=rsbody("ggxh")%></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("yasjgl")%>&nbsp;</td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("zlfs")%></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ccbh")%>&nbsp;</td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("changj")%></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</table>


<%call showpage(page,url,total,record,PgSz)
end sub


sub pdc()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td  style="border-bottom-style: solid;border-width:1px" width="3%"><div align="center"><strong>���      </strong></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="13%"><div align="center"><strong>λ��</strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="7%"><div align="center"><strong>λ ��</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ͺ�</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>���ش������ͺ�</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ٶȴ������ͺ�</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>������Χ</strong></div></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="5%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ע</strong></div></td>
      <td width="15%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѡ ��</strong></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=13 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
  
	xh=xh+1%>
           

		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td  style="border-bottom-style: solid;border-width:1px" width="3%"><div align="center"><%=(xh)%></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="13%"><%=rsbody("llname")%></td>
      <td style="border-bottom-style: solid;border-width:1px" width="7%"><%=rsbody("wh")%></td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ggxh")%>&nbsp;</td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("pdcczcgqxh")%></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("pdcsdcgqxh")%></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("clfw")%>&nbsp;</td>
      <td width="8%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("gyjz")%></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("changj")%></td>
      <td width="5%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</table>


<%call showpage(page,url,total,record,PgSz)
end sub


sub tjf()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td  style="border-bottom-style: solid;border-width:1px" width="6%"><div align="center"><strong>���      </strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="12%"><div align="center"><strong>λ ��</strong></div></td>
      <td width="24%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>����</strong></div></td>
      <td width="31%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ͺ�</strong></div></td>
      <td width="27%" style="border-bottom-style: solid;border-width:1px"><div align="center"></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=14 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
  
	xh=xh+1%>
           

		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td  style="border-bottom-style: solid;border-width:1px" width="6%"><div align="center"><%=(xh)%></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="12%"><%=rsbody("wh")%>&nbsp;</td>
      <td width="24%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("llname")%>&nbsp;</td>
      <td width="31%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ggxh")%>&nbsp;</td>
      <td width="27%" style="border-bottom-style: solid;border-width:1px"><div align="center"><a href="ylb_view.asp?id=<%=rsbody("id")%>">����鿴��ϸ����</a></div></td>
    </tr>
    <%RowCount=RowCount-1

	  rsbody.movenext
        loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</table>


<%call showpage(page,url,total,record,PgSz)
end sub


sub ddzxjg()
 
%>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td  style="border-bottom-style: solid;border-width:1px" width="5%"><div align="center"><strong>���      </strong></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="6%"><div align="center"><strong>����</strong></div></td>
      <td style="border-bottom-style: solid;border-width:1px" width="14%"><div align="center"><strong>λ ��</strong></div></td>
      <td width="17%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�ͺ�</strong></div></td>
      <td width="13%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ʽ</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ��</strong></div></td>
      <td width="7%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>�� ע</strong></div></td>
      <td width="17%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>ѡ ��</strong></div></td>
    </tr>
		  <%sqlbody="SELECT * from ylbbody where dclass=15 order by  id DESC"
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>��������</p>" 
    else
      
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
dim start
start=PgSz*Page-PgSz+1
   rowCount = rsbody.PageSize
  do while not rsbody.eof and rowcount>0
    xh=xh+1%>
           

		   
		   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td  style="border-bottom-style: solid;border-width:1px" width="5%"><div align="center"><%=(xh)%></div></td>
      <td  style="border-bottom-style: solid;border-width:1px" width="6%"><%=(sscjh_d(rsbody("sscj")))%></td>
      <td style="border-bottom-style: solid;border-width:1px" width="14%"><%=rsbody("wh")%></td>
      <td width="17%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ggxh")%>&nbsp;</td>
      <td width="13%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ddzxjg_xs")%></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("ddzxjg_fk")%></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("changj")%></td>
      <td width="7%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("whbeizhu")%>&nbsp;</div></td>
      <%response.write "<td width=""17%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center"">  <a href=ylb_jxjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;<a href=ylb_ghjl.asp?ylbid="&rsbody("id")&"&lxclassid="&lxclassid&">����</a>&nbsp;"
	      	        call editdel(rsbody("id"),rsbody("sscj"),"ylb_ned.asp?action=edit&lxclassid="&lxclassid&"&id=","ylb_ned.asp?action=del&id=")'���ޡ��������༭��ɾ��

response.write "</div></td></tr>"
 RowCount=RowCount-1
   rsbody.movenext
  loop
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
response.write "</table>"
call showpage(page,url,total,record,PgSz)
end sub

sub search(lxclassid)
dim rscj,sqlcj

response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
response.write "<form method='Get' name='SearchForm' action='ylb_search.asp' onsubmit='javascript:return CheckSearch();'>" & vbCrLf
response.write "  <tr class='tdbg'>   <td>" & vbCrLf
response.write "  <strong>λ��������</strong>" & vbCrLf

	response.write"<select name='lxclassid' size='1' >"& vbCrLf
    response.write"<option  selected>ѡ���豸����</option>"& vbCrLf
	sqlcj="SELECT * from lxclass where lxznum=0 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("lxnum")&"'>"&rscj("lxname")&"</option>"& vbCrLf
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select>"  	 & vbCrLf
response.write "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"">" & vbCrLf
response.write "  <input type='submit' name='Submit'  value='����'>" & vbCrLf
response.write "  <input type='hidden' name='action' value='keys'>" & vbCrLf
response.write "</td></form><td width='50%'><strong>�鿴���������������ݣ�</strong>"

response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "<option value=''>��������ת����</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
        response.write"<option value='ylb_search.asp?action=sscjs&lxclassid="&lxclassid&"&sscj="&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	response.write "     </select>	" & vbCrLf
response.write "</tr></table>" & vbCrLf


end sub
Call CloseConn
%>