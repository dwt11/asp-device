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
dim start,url,key,pagename
dim sqllsda,rslsda,ii,lxclassid,tyzk,zxzz
dim record,pgsz,total,page,rowCount,xh,sscj
action = Trim(Request("action"))
key=trim(request("keyword"))
sscj=trim(request("sscj"))
ssgh=trim(request("ssgh"))

if action="keys" then 
  pagename="�ؼ���"""&key&"&nbsp;""�� �������� �е��������"
  url="lsda_search.asp?keyword="&key&"&action=keys"
end if   
if action="sscjs" then 
   pagename=sscjh(sscj)&" ���� ��������"
   url="lsda_search.asp?sscj="&sscj&"&action=sscjs"
end if 
if action="ssghs" then 
   pagename=gh(ssgh)&" ���� ��������"
   url="lsda_search.asp?ssgh="&ssgh&"&action=ssghs"
end if 

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>��Ϣ����ϵͳ������������ҳ</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'  onload='redirect("
response.write sscj&","&ssgh
response.write")'>"& vbCrLf
response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>"&pagename&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "<tr class='tdbg'>"& vbCrLf
response.write "    <td width='70' height='30'><strong>��������</strong></td>"& vbCrLf
response.write "    <td height='30'><a href=""lsda.asp"">����������ҳ</a>&nbsp;|&nbsp;<a href=""lsda.asp?action=add"">�����������</a>&nbsp;|&nbsp;<a href=""tocsv.asp?action=lsdamain&titlename=��������"" target=""_blank"">���������������̨�˵�Excel�ĵ�</a>  ��·<font color='ff0000'>��ɫ</font>Ϊԭ��<font color='0000ff'>��ɫ</font>Ϊ����ԭ��</td>"& vbCrLf
response.write "  </tr>"& vbCrLf

response.write "</table>"& vbCrLf
call search()
dim v1 ,v2,v3,totall
v1= "<font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=1")(0)&"/"& Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=1")(0) 
v1= v1&"&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=1")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=1")(0)*100,5)&"%</font>" 
v1=v1&"&nbsp;"&"<font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=1 and czyy=0")(0)&"</font>/<font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=1 and czyy")(0)&"</font>"

v2= "<font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=2")(0)&"/"& Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=2")(0) 
v2= v2&"&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=2")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=2")(0)*100,5)&"%</font>" 
v2=v2&"&nbsp;<font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=2 and czyy=0")(0)&"</font>/<font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=2 and czyy")(0)&"</font>"

v3= "<font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=3")(0)&"/"& Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=3")(0) 
v3= v3&"&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1 and sscj=3")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE sscj=3")(0)*100,5)&"%</font>" 
v3=v3&"&nbsp;<font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=3 and czyy=0")(0)&"</font>/<font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=0 and sscj=3 and czyy")(0)&"</font>"


totall= "<font color='#006600'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1")(0)&"/"& Connjg.Execute("SELECT count(lsdaid) FROM lsda")(0) 
totall= totall&"&nbsp;"&left(Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE tyzk=1")(0)/Connjg.Execute("SELECT count(lsdaid) FROM lsda")(0)*100,5)&"%</font>" 
totall=totall&"&nbsp;<font color='#0000ff'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE  tyzk=0 and czyy=0")(0)&"</font>/<font color='#ff0000'>"&Connjg.Execute("SELECT count(lsdaid) FROM lsda where  tyzk=0 and czyy")(0)&"</font>" 

response.write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"& vbCrLf
response.write " <tr >"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>άһͶ���ʣ�"&v1&"</strong>   <strong>ά��Ͷ���ʣ�"&v2&"</strong>     <strong>ά��Ͷ���ʣ�"&v3&"</strong> <br>    <strong>��Ͷ���ʣ�"&totall&"</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write "</table>"& vbCrLf
if action="keys" then sqllsda="SELECT * from lsda where wh like '%" & key & "%' ORDER BY SSCJ ASC,ssGH ASC,WH ASC"
if action="sscjs" then sqllsda="SELECT * from lsda where sscj="&sscj&" ORDER BY SSCJ ASC,ssGH ASC,WH ASC"
if action="ssghs" then sqllsda="SELECT * from lsda where ssgh="&ssgh&" ORDER BY SSCJ ASC,ssGH ASC,WH ASC"
    set rslsda=server.createobject("adodb.recordset")
    rslsda.open sqllsda,connjg,1,1
    if rslsda.eof and rslsda.bof then 
      message("δ�������������") 
    else
      record=rslsda.recordcount
      if Trim(Request("PgSz"))="" then
         PgSz=20
      ELSE 
         PgSz=Trim(Request("PgSz"))
      end if 
      rslsda.PageSize = Cint(PgSz) 
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
      rslsda.absolutePage = page
      start=PgSz*Page-PgSz+1
      rowCount = rslsda.PageSize
      response.write "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      response.write "<tr class=""title"">"
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>���</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>����</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>װ��</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>λ��</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��;</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>һ�μ�����</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��λ</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��Χ</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ֵL</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>����ֵH</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>Ͷ��״��</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ִ��װ��</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>�ּ�</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>��ע</strong></div></td>"
response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>ѡ��</strong></div></td>"
      response.write "</tr>"
      do while not rslsda.eof and rowcount>0
	    'xh=xh+1
                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
                response.write "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rslsda("lsdaid")&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sscjh_d(rslsda("sscj"))&"</div></td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&gh(rslsda("ssgh"))&"</div></td>"
					 'hbody=searchH(cstr(key),cstr(rslsda("wh")))
           response.write "      <td style=""border-bottom-style: solid;border-width:1px"""
		   if now()-rslsda("update")<7 then   response.write "bgcolor=""#FFFF00"""
		   
		   
		   response.write ">"&searchH(rslsda("wh"),key)&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("yt")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("ycjname")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("cldw")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("clfw")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("lsl")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("lsh")&"&nbsp;</td>"
         select case rslsda("tyzk")
          case 0
             tyzk="��·"
			 if rslsda("czyy") then
		        tyzk="<font color='#ff0000'>"&tyzk&"</font>"
		      else
		        tyzk="<font color='#0000ff'>"&tyzk&"</font>"
		     end if 	
          case 1 
        	tyzk="<font color='#006600'>Ͷ��</font>"
          
        end select	 
				response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href='lsda_czjl.asp?lsdaid="&rslsda("lsdaid")&"&lsdawh="&rslsda("wh")&"'>"&tyzk&"</a>&nbsp;</div></td>"
                zxzz=rslsda("zxzz")
				if len(zxzz)>7 then 
				  zxzz=left(zxzz,6)&"��"
                   
				      response.write"<script language=javascript src='/js/showPopupText.js'></script>"
                      response.write "      <td style=""border-bottom-style: solid;border-width:1px"" onmouseover=""pop('"&rslsda("zxzz")&"','#3366CC');"">"&zxzz&"&nbsp;</td>"
                else
				  response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&zxzz&"&nbsp;</td>"
				end if 
			    response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&fj(rslsda("fj"))&"</td>"
			    response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"&rslsda("bz")&"&nbsp;</td>"
                response.write "      <td style=""border-bottom-style: solid;border-width:1px"">"
				call editdel_d(rslsda("lsdaid"),rslsda("sscj"),"lsda.asp?action=edit&id=","lsda.asp?action=del&id=")
				
                response.write "</td></tr>"
        RowCount=RowCount-1
        rslsda.movenext
      loop
      response.write "</table>"
      call showpage(page,url,total,record,PgSz)
   end if
   rslsda.close
   set rslsda=nothing
   connjg.close
   set connjg=nothing


response.write "</body></html>"

sub search()
dim sqlcj,rscj

response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
response.write "<form method='Get' name='SearchForm' action='lsda_search.asp'>" & vbCrLf
response.write "  <tr class='tdbg'>   <td>" & vbCrLf
response.write "  <strong>λ��������</strong>" & vbCrLf
response.write "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"" value="&key&">" & vbCrLf
response.write "  <input type='Submit' name='Submit'  value='����'>" & vbCrLf
response.write "  <input type='hidden' name='Action' value='keys'>" & vbCrLf
response.write "</td></form><td><font color='0066CC'> �����������ݣ�</font>"
response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "	       <option value=''>��������ת����</option>" & vbCrLf
sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='lsda_search.asp?action=sscjs&sscj="&rscj("levelid")&"'"
		if cint(request("sscj"))=rscj("levelid") then response.write" selected"
		response.write">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	response.write "     </select>	" & vbCrLf
	
	
	
response.write "	<font color='0066CC'> ����װ�����ݣ�</font>"
response.write "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
response.write "	       <option value=''>��װ����ת����</option>" & vbCrLf
sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
       	response.write"<option value='lsda_search.asp?action=ssghs&ssgh="&rsgh("ghid")&"'"
		if cint(request("ssgh"))=rsgh("ghid") then response.write" selected"
		response.write">"&rsgh("gh_name")&"("&Connjg.Execute("SELECT count(lsdaid) FROM lsda WHERE ssgh="&rsgh("ghid"))(0)&")</option>"& vbCrLf
	
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	response.write "     </select>	" & vbCrLf
	
	
'	
'	   response.write"<select name='lsda_sscj' size='1'  onChange=""redirect(this.options.selectedIndex,0)"">"& vbCrLf
'    response.write"<option  selected>ѡ����������</option>"& vbCrLf
'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    do while not rscj.eof
'       	response.write"<option value='"&rscj("levelid")&"' "
'		if cint(request("sscj"))=rscj("levelid") then response.write" selected"
'
'		response.write ">"&rscj("levelname")&"</option>"& vbCrLf
'	
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'    response.write"</select>"  	 & vbCrLf
'    'response.write "<select name='lsda_gh' size='1'  onChange=""alert(document.all.lsda_sscj.options[document.all.lsda_sscj.selectedIndex].value);alert(this.value);"">" & vbCrLf
'	response.write "<select name='lsda_gh' size='1' onChange=""location.href='lsda_search.asp?action=sscjs&sscj=' + document.all.lsda_sscj.options[document.all.lsda_sscj.selectedIndex].value + '&ssgh=' + this.value;"">" & vbCrLf
'    response.write "<option  selected>ѡ��װ�÷���</option>" & vbCrLf
'    response.write "</select></td></tr>  "  & vbCrLf
'    response.write "<script><!--" & vbCrLf
'    response.write "var groups=document.all.lsda_sscj.options.length" & vbCrLf
'    response.write "var group=new Array(groups)" & vbCrLf
'    response.write "for (i=0; i<groups; i++)" & vbCrLf
'    response.write "group[i]=new Array()" & vbCrLf
'    response.write "group[0][0]=new Option(""ѡ��װ�÷���"","" "");" & vbCrLf
'	
'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    if rscj.eof then 
'	  response.write "δ�ҵ�����"
'    else
'	do while not rscj.eof
'     lsdaii=1		
'		sqlgh="SELECT * from ghname where sscj="&rscj("levelid")
'        set rsgh=server.createobject("adodb.recordset")
'        rsgh.open sqlgh,conn,1,1
'        if rsgh.eof and rsgh.bof then
'		   response.write "group["&rscj("levelid")&"][0]=new Option(""δ���װ��"",""0"");" & vbCrLf
'		else
'		   response.write"group["&rsgh("sscj")&"][0]=new Option(""ѡ��װ�÷���"",""0"");" & vbCrLf
'		do while not rsgh.eof
'		   response.write"group["&rsgh("sscj")&"]["&lsdaii&"]=new Option("""&rsgh("gh_name")&""","""&rsgh("ghid")&""");" & vbCrLf
'		  lsdaii=lsdaii+1
'		   rsgh.movenext
'	    loop
'	    end if 
'		rsgh.close
'	    set rsgh=nothing
'
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'
'  end if 
'
'
'    response.write "var temp=document.all.lsda_gh" & vbCrLf
'    response.write "function redirect(x,y){" & vbCrLf
'    response.write "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
'    response.write "temp.options[m]=null" & vbCrLf
'    response.write "for (i=0;i<group[x].length;i++){" & vbCrLf
'    
'	
'	response.write "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
'    response.write "if(group[x][i].text= y ){"
'	response.write "temp.options[i].selected=true;"
'	response.write "alert(group[x][i].text);}"
'  
'  
'  
'  '�˴�group[x][i].text  /.value ��������ֵһֱΪ1 �����ж�  07  29
'  
'  
'  
'  
'    response.write "}" & vbCrLf
'    'response.write "temp.options[y].selected=true" & vbCrLf
'    'response.write "location.href=""lsda_search.asp?action=sscjs&sscj=""+x + ""&ssgh="" + group[x][0].value"
'	response.write "}//-->" & vbCrLf  ȱ��JS������־
response.write "	</td>  </tr></table>" & vbCrLf


end sub
	
Call Closeconn
%>