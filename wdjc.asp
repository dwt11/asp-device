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
'���ݿ��� txdate�ֶ�Ϊ�û���ѡֵ��ʱ�䣬txdate1Ϊʵ����д��ʱ�䣬Ĭ������
dim sqlzblog,rszblog,title,record,pgsz,total,page,start,rowcount,xh,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>��Ϣ����ϵͳ--�����豸Ѳ���¼</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/tab.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
	call saveadd
  case "del"
    'if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
	call del
  case ""
	'if truepagelevelh(session("groupid"),0,session("pagelevelid")) then 
	call main
end select	

sub add()
  	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:20px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>���Ѳ���¼</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='wdjc.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
	
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

    if session("level")=3 then 
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >��������:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element>"& vbCrLf
		dwt.out"<input class='x-form-text x-form-field' style='WIDTH: 175px' value='"&ssbzh(session("levelzclass"))&"'  disabled='disabled' >"& vbCrLf
		dwt.out"<input name='ssbz' type='hidden' value="&session("levelclass")&">"& vbCrLf
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	end if 
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>�����:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=jxjl_jxrname  disabled='disabled' value='"&conn.Execute("SELECT username1 FROM userid WHERE id="&session("userid"))(0)&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>ʱ��:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element'>"& vbCrLf
    dwt.out"<input name='txdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	
	
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 75px'>Ѳ���:</LABEL>"& vbCrLf
		dwt.out"				<DIV class='x-form-element'>"& vbCrLf
	  sqlclass="SELECT * from address where ssbz="&session("levelzclass")&" order by id"
	  set rsclass=server.createobject("adodb.recordset")
	  rsclass.open sqlclass,connw,1,1
	  if rsclass.eof and rsclass.bof then 
		  dwt.out  message ("<p align='center'>δ���Ѳ���豸</p>" )
	  else
	  
	  
	  
		  dwt.out "<table   border='1'  cellpadding='1' cellspacing='1'>"
	  	  dwt.out "<tr>"
		  
		  dwt.out "<td align=center>���</td>"
	  	  dwt.out "<td align=center>λ��</td>"
	  	  dwt.out "<td align=center>����</td>"
	  	  dwt.out "<td align=center>�¶�</td>"
		  dwt.out "</tr>"
	  do while not rsclass.eof 
	  
	  
	  	  dwt.out "<tr>"
		  
		  dwt.out "<td>"&rsclass("id")&"</td><input name='bh"&rsclass("id")&"' type='hidden' value='"&rsclass("id")&"'>"
		  
	  	  dwt.out "<td>"&rsclass("wz")&"</td><input name='wz"&rsclass("id")&"' type='hidden' value='"&rsclass("wz")&"'>"
	  	  dwt.out "<td>"&rsclass("name")&"</td><input name='name"&rsclass("id")&"' type='hidden' value='"&rsclass("name")&"'>"
	  	  dwt.out "<td><input name='ti"&rsclass("id")&"' type='text' ></td>"
		  dwt.out "</tr>"
	  
	  rsclass.movenext
	  loop
	  dwt.out "</table>"
	  end if 

		
		
		
		
		
		
		
		
		
		
		
dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf


	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""location.href='';"" style='cursor:hand;'>"& vbCrLf
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
	dwt.out"</DIV>"& vbCrLf
	
	
end sub	

sub saveadd()    
	 
	  '����
      
	  
	  
	  
	  sqlclass="SELECT * from address"
	  set rsclass=server.createobject("adodb.recordset")
	  rsclass.open sqlclass,connw,1,1
	  if rsclass.eof and rsclass.bof then 
		  'dwt.out  message ("<p align='center'>δ���Ѳ���豸</p>" )
	  else
	  do while not rsclass.eof 
		  
		  if request("wz"&rsclass("id"))<>"" and  request("name"&rsclass("id"))<>"" and request("ti"&rsclass("id"))<>"" then
		  
			
			set rsadd=server.createobject("adodb.recordset")
			sqladd="select * from bb" 
			rsadd.open sqladd,connw,1,3
			rsadd.addnew
			rsadd("userid")=session("userid")
			rsadd("ssbz")=session("levelzclass")
			rsadd("wz")=request("wz"&rsclass("id"))
			rsadd("name")=request("name"&rsclass("id"))
			rsadd("ti")=request("ti"&rsclass("id"))
			rsadd("update")=now()
			rsadd.update
			rsadd.close
			set rsadd=nothing
		  end if 
	  rsclass.movenext
	  loop
	  end if
	  
	  
	  dwt.savesl "�¶�Ѳ���¼","���",now()
	 
	 
		
		
		

	 dwt.out "<Script Language=Javascript>location.href='?ssbz="&session("levelzclass")&"';</Script>"
end sub



sub main()
	url=GetUrl
	dwt.out "<div style='left:6px;'>"
	dwt.out "     <DIV class='x-layout-panel-hd x-layout-title-center'>"
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><b>�����豸����¼</b></span>"
	dwt.out "     </div>"
	dwt.out "</div>"

	dwt.out "<div class='x-toolbar' style='padding-left:15px;'>"
	dwt.out "	<div align=left>"
    	dwt.out "		 <a href='?action=add'>��Ӽ�¼</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='wdjc_class.asp'>Ѳ���豸����</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='wdjc_class.asp?action=add'>���Ѳ���豸</a>"
'	
	
	dwt.out "	ʹ��˵�����԰������û�����¼�󣬵�������Ѳ���豸���������ɺ�������Ӽ�¼���ڶ�Ӧ���豸����д�¶ȼ��ɡ�Ѳ���豸ֻ��Ҫ���һ�μ���</div>"
	dwt.out "</div>"

   
   
   
	

	dwt.out "<div class='navg'>"
	dwt.out "  <div id='system' class='mainNavg'>"
	dwt.out "    <ul>"
		sscjid=request("sscj")
		'101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
		 
		 
		  if sscjid="" then 
		     if  request("ssbz")="" then 
    			  sscjid=1    '101218�޸ģ���ҳ����Զ���ʾ��Ӧ�ĳ���
			 else
					sqlbz="SELECT * from bzname where id="&request("ssbz")
					set rsbz=server.createobject("adodb.recordset")
					rsbz.open sqlbz,conn,1,1
					if rsbz.eof and rsbz.bof then 
						'dwt.out  message ("<p align='center'>δ��Ӱ���</p>" )
sscjid=0
					else
						sscjid=rsbz("sscj")
					end if 

			 end if 	  
			  
		  end if 	  
			  
	
	
	
	sqlsscj="SELECT * from levelname where levelclass=1 and levelid<4"
	set rssscj=server.createobject("adodb.recordset")
	rssscj.open sqlsscj,conn,1,1
	if rssscj.eof and rssscj.bof then 
		dwt.out  message ("<p align='center'>δ�����������</p>" )
	else
	do while not rssscj.eof 
		if cint(sscjid)=rssscj("levelid") then 
		   dwt.out "<li id='systemNavg'><a href='#'>"&rssscj("levelname")&"</a></li>"
		else
		   dwt.out "<li><a href='?sscj="&rssscj("levelid")&"'>"&rssscj("levelname")&"</a></li>"
		end if    
	rssscj.movenext
	loop
	end if 
	  
	  
    dwt.out "</ul>"
    dwt.out " </div>"
	
	dwt.out "  <div class='textbody'>"
		sqlssbz="SELECT * from bzname where sscj="&sscjid
		set rsssbz=server.createobject("adodb.recordset")
		rsssbz.open sqlssbz,conn,1,1
		if rsssbz.eof and rsssbz.bof then 
			'dwt.out  message ("<p align='center'>δ��Ӱ���</p>" )
		else
		
		
		
		
		do while not rsssbz.eof 



	  dwt.out "<b><a href='?ssbz="&rsssbz("id")&"'>"&rsssbz("bzname")&"</a></b> "
	  ij=ij+1
	  if ij=1 then 
		ssbzid=rsssbz("id")
		ssbzname=rsssbz("bzname")
	  end if 
	  
	rsssbz.movenext
	loop
	end if 
	
	  if request("ssbz")<>"" then 
		ssbzid=request("ssbz")
			  sqlbz="SELECT * from bzname where id="&ssbzid
			  set rsbz=server.createobject("adodb.recordset")
			  rsbz.open sqlbz,conn,1,1
			  if rsbz.eof and rsbz.bof then 
				 ' dwt.out  message ("<p align='center'>δ��Ӱ���</p>" )
			  else
				  ssbzname=rsbz("bzname")
			  end if 
	  end if 
	
	
	


		  dwt.out "<br><table  border='1'  cellpadding='1' cellspacing='1'>"
		  dwt.out "		  <tr>"
		  dwt.out "		    <td colspan='5' align=center><b>"&ssbzname&"</b></td>"
		  dwt.out "	      </tr>"
			  dwt.out "		  <tr>"
			  dwt.out "		    <td align=center>����</td>"
			  dwt.out "		    <td align=center>λ��</td>"
			  dwt.out "		    <td align=center>����</td>"
			  dwt.out "		    <td align=center>�¶�</td>"
			  dwt.out "		    <td>&nbsp;</td>"
			  dwt.out "	      </tr>"



		  sqljl="SELECT * from bb where ssbz="&ssbzid&" order by wz,update desc" 
		  set rsjl=server.createobject("adodb.recordset")
		  rsjl.open sqljl,connw,1,1
		  if rsjl.eof and rsjl.bof then 
		  dwt.out "		  <tr>"
		  dwt.out "		    <td colspan='5' align=center>δ��Ӽ�¼</td>"
		  dwt.out "	      </tr>"
		  else
		  record=rsjl.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
	   end if 
	   rsjl.PageSize = Cint(PgSz) 
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
	   rsjl.absolutePage = page
	   start=PgSz*Page-PgSz+1
	   rowCount = rsjl.PageSize
	   do while not rsjl.eof and rowcount>0
			 ' dwt.out rsjl("update")&" "&rsjl("wz")&" "&rsjl("ti")&"<br>"
			  dwt.out "		  <tr>"
			  dwt.out "		    <td>"&rsjl("update")&"</td>"
			  dwt.out "		    <td>"&rsjl("wz")&"</td>"
			  dwt.out "		    <td>"&rsjl("name")&"</td>"
			  dwt.out "		    <td>"&rsjl("ti")&"</td>"
			  dwt.out "		    <td><a href=wdjc_view.asp?name="&Server.URLEncode(rsjl("name"))&"&ssbz="&rsjl("ssbz")&"&wz="&Server.URLEncode(rsjl("wz"))&"  target='_blank'>����λ�����м�¼</a>   "
				if session("level")=1 or session("level")=0 and session("levelclass")=sscj then  response.write "<a href=?action=del&id="&rsjl("id")&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a>"
			
			  dwt.out "</td>"
			  dwt.out "	      </tr>"
		
		RowCount=RowCount-1
		  rsjl.movenext
		  loop
		  end if 
	
		  dwt.out "</table><br>"
	
if request("ssbz")<>"" or request("sscj")<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 	
	
	
	
	
	
	
	
	dwt.out "</div>"
	dwt.out "</div>	"
	
	
	rsjl.close
	set rsjl=nothing
	conn.close
	set conn=nothing
end sub	

	
	
	
	
	
	
	
	

sub del()
ID=request("ID")



set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from bb where id="&id
rsdel.open sqldel,connw,1,3
dwt.out "<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  
dwt.savesl "�¶�Ѳ���¼","ɾ��",now()

end sub






dwt.out  "</body></html>"

Call CloseConn
%>
