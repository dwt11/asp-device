<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'dim starttime : starttime=timer
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<%
if request("whl")=2 then title="-������豸�б�"
if request("tyl")=2 then title="-δͶ���豸�б�"
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>������������ҳ</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/grid.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
	dwt.out "<div style='left:26px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�豸���� ��ҳ"&title&"</span>"
	dwt.out "     </div>"
	dwt.out "<div class='x-toolbar' ><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='sb_total.asp'>" & vbCrLf
	dwt.out "<input name='action' type='hidden' id='action' value='search'>"
	dwt.out "&nbsp;&nbsp;<input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 dwt.out "value='"&request("keyword")&"'"
		dwt.out ">" & vbCrLf
	else
	 dwt.out "value='����������λ��'"
	 dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	dwt.out "  <input type='submit' name='Submit'  value='����'>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "��ʾ:ֻ������λ�ŵĲ������ݼ�������������</form></div></div>" & vbCrLf
action=request("action")
select case action
  case ""
      call main
  case "search"
      call search
end select	  	 


'ȡ�ӷ�������
function zclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclass= "δ�༭" 
  else
     zclass=rsbody("sbclass_name")
  end if
end function

sub search()
url=geturl
		keys=request("keyword")
   if keys<>"" then 
		sqlbody="SELECT * from sb where sb_wh  like '%" &keys& "%' order by sb_dclass,sb_sscj aSC,sb_ssgh asc,sb_wh asc"
   end if 	
   if request("whl")=2 then 
		'keys=request("keyword")
		sqlbody="SELECT * from sb where sb_whqk=2 and sb_sscj="&request("sscj")&" order by sb_dclass,sb_ssgh asc,sb_wh asc"
   end if 	
   if request("tyl")=2 then 
		'keys=request("keyword")
		sqlbody="SELECT * from sb where (sb_tyqk=2 or sb_tyqk=3 ) and sb_sscj="&request("sscj")&" order by sb_dclass,sb_ssgh asc,sb_wh asc"
   end if 	
	set rsbody=server.createobject("adodb.recordset")
	rsbody.open sqlbody,conn,1,1
	if rsbody.eof and rsbody.bof then 
		message "<p align=""center"">δ�ҵ��������</p>" & vbCrLf
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>���</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>λ��</div></td>" & vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>" & vbCrLf
dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>װ��</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>ѡ��</div></td>" & vbCrLf
		dwt.out "    </tr>" & vbCrLf
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
		do while not rsbody.eof  and rowcount>0
         			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
					dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='sb.asp?sbclassid="&rsbody("sb_dclass")&"&keyword="&rsbody("sb_wh")&"'>"&searchH(uCase(rsbody("sb_wh")),keys)&"</a></div></td>" & vbCrLf
					dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsbody("sb_sscj"))&"</div></td>" & vbCrLf
					dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""left"">"&GH(rsbody("sb_ssGH"))&"</div></td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&zclass(rsbody("sb_dclass"))&"-"&zclass(rsbody("sb_zclass"))&"</td>" & vbCrLf
					dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=sb_jxjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&rsbody("sb_dclass")&">����</a>&nbsp;<a href=sb_ghjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&rsbody("sb_dclass")&">����</a></div>" & vbCrLf

		   'dwt.out sscjh(rsbody("sb_sscj"))&" "&zclass(rsbody("sb_dclass"))&"-"&zclass(rsbody("sb_zclass"))&" "&searchH(uCase(rsbody("sb_wh")),keys)&"<br/>"
		    dwt.out "</tr>"
		RowCount=RowCount-1
		rsbody.movenext
		loop
		dwt.out "</table>"
		call showpage(page,url,total,record,PgSz)
		dwt.out "</div>"
	end if 
	dwt.out "</div>"
    rsbody.close
	set conn=nothing
	
end sub
sub main()
	
	dwt.out "<table  width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
	dwt.out "<tr class=""x-grid-header"">" 
	dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>�����</div></td>"
	'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>׼ȷ��</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>Ͷ����</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'></div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'></div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'></div></td>"
	dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'></div></td>"
	dwt.out "    </tr>"
	dwt.out "<tr class=""title"">" 
	
	dwt.out "    </tr>"
		dim sqlcj,rscj
		sqlcj="SELECT * from levelname where levelclass=1 and levelid<5 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
	dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" 
	
	
	'����
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rscj("levelname")&"</div></td>"
	whqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_whqk=1 and sb_sscj="&rscj("levelid")&"")(0)
	'zqqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_zqqk=3 and sb_sscj="&rscj("levelid")&"")(0) 
	tyqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_tyqk=1 and sb_sscj="&rscj("levelid")&"")(0) 
	total_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_sscj="&rscj("levelid")&"")(0)
	
	wh_l=left(whqk_numb/total_numb,5)*100&"%"
	zq_l=left(zqqk_numb/total_numb,5)*100&"%"
	ty_l=left(tyqk_numb/total_numb,5)*100&"%"
	
	'���
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><span style=""color:'#006600'""><a href=sb_total.asp?action=search&whl=2&sscj="&rscj("levelid")&">"&wh_l&"("&whqk_numb&")</a></span></div></td>"
	
	'׼ȷ
	'dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><span style=""color:'#006600'"">"&zq_l&"("&zqqk_numb&")</span></div></td>"
	
	'Ͷ��
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><span style=""color:'#006600'""><a href=sb_total.asp?action=search&tyl=2&sscj="&rscj("levelid")&">"&ty_l&"("&tyqk_numb&")</a></span></div></td>"
	
	'����
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&total_numb&"</span></div></td>"
	
	
	dwt.out "    </tr>"
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		
	dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" 
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""></div></td>"
	
	
	whqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_whqk=1 AND sb_sscj<5")(0)
	'zqqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_zqqk=3 AND sb_sscj<5")(0) 
	tyqk_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_tyqk=1")(0) 
	total_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_sscj=1 or sb_sscj=2 or sb_sscj=3 or sb_sscj=4 ")(0)
	
	wh_l=left(whqk_numb/total_numb,5)*100&"%"
	zq_l=left(zqqk_numb/total_numb,5)*100&"%"
	ty_l=left(tyqk_numb/total_numb,5)*100&"%"
	
	
	
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><span style=""color:'#006600'"">"&wh_l&"("&whqk_numb&")</span></div></td>"
	'dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><span style=""color:'#006600'"">"&zq_l&"("&zqqk_numb&")</span></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><span style=""color:'#006600'"">"&ty_l&"("&tyqk_numb&")</span></div></td>"
	dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&total_numb&"</span></div></td>"
	dwt.out "    </tr>"
	dwt.out"</table>"
end sub
	dwt.out "</div>"
dwt.out "</body></html>"

%>