<%@language=vbscript codepage=936 %>
<%
'Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/sb_function.asp"-->
<%dim url
dim sqlbody,rsbody,ii,sb_classid,sbid,sqlgh,rsgh
dim record,pgsz,total,page,rowCount,xh,sb_sscj
dim sb_wh,sql,rs

url=geturl
keys=ReplaceBadChar(trim(request("keyword")) )
sscjid=trim(request("sscj")) 
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>������¼����</title>"& vbCrLf
dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"</head>"& vbCrLf
dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>������¼����</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf

    call search()
 if request("action")="" or request("action")="dcs" then call main
if request("action")="ghcount" then call ghcount

sub main()



  if request("action")="dcs" then  sqljxwhere = " and (sb.sb_dclass=121 or sb.sb_dclass=105 or sb.sb_dclass=144 or sb.sb_dclass=123 or sb.sb_dclass=124 or sb.sb_dclass=125 or sb.sb_dclass=127 or sb.sb_dclass=128 or sb.sb_dclass=129 or sb.sb_dclass=133 or sb.sb_dclass=134 or sb.sb_dclass=126 or sb.sb_dclass=84 or sb.sb_dclass=88 or sb.sb_dclass=91 or sb.sb_dclass=94 or sb.sb_dclass=95 or sb.sb_dclass=96 or sb.sb_dclass=97 or sb.sb_dclass=98 or sb.sb_dclass=99 )"
  if sscjid<>"" then sqljxwhere=" and (((sb.sb_sscj)="&sscjid&")) "
  if keys<>"" then sqljxwhere=" and (sb.sb_wh like '%" &keys& "%'  "
  if keys<>"" then sqljxwhere1=" or  jx_ren like '%" &keys& "%'  or jx_fzren like '%" &keys& "%'  ) "
  if keys<>"" then sqljxwhere2=" or  gh_ren like '%" &keys& "%'  ) "   '�˶�����������  �����ľ�����
 
 
  'sqljx=sqljx&"order by  gh_date DESC"	


sqljx="SELECT sbjx.*,sbgh.gh_xh as gh_xh,sbgh.gh_xhupdate as gh_xhupdate,sb.sb_sscj,sb.sb_wh,sb.sb_dclass from"
sqljx=sqljx&" (sbjx as sbjx left join sbgh as sbgh on sbjx.jx_id=sbgh.jx_id) "
sqljx=sqljx&"   left JOIN sb ON sbjx.sb_id = sb.sb_id "
sqljx=sqljx&"  where instr(sbjx.jx_nr_new,'9999')>0 " & sqljxwhere & sqljxwhere1 & " order by  sbjx.jx_DATE DESC"
'sqljx=sqljx&"  UNION ALL select null as jx_id,sbgha.sb_id as sb_id,null as jx_name,null as jx_lb,sbgha.gh_yy as jx_gzxx,null as jx_gzxx_new,null as jx_nr,null as jx_nr_new,sbgha.gh_date as jx_date,null as jx_enddate,null as jx_fzren,sbgha.gh_ren as jx_ren,null as jx_ylwt,sbgha.gh_bz as jx_bz,sbgha.gh_xh as gh_xh, sbgha.gh_xhupdate as gh_xhupdate,sb.sb_sscj,sb.sb_wh,sb.sb_dclass from"
'sqljx=sqljx&"  sbgh as sbgha  left JOIN sb ON sbgha.sb_id = sb.sb_id where 1=1  " & sqljxwhere & sqljxwhere2 & "  order by jx_date  desc "
  
  'sqljx=sqljx&"  order by jx_date  desc "
'sqljx="  SELECT sbjx.*,sbgh.gh_xh as gh_xh,sbgh.gh_xhupdate as gh_xhupdate,sb.sb_sscj,sb.sb_wh,sb.sb_dclass from"
'sqljx=sqljx&" (sbjx as sbjx left join sbgh as sbgh on sbjx.jx_id=sbgh.jx_id) "
'sqljx=sqljx&"  left JOIN sb ON sbgh.sb_id = sb.sb_id "
'sqljx=sqljx&" where instr(sbjx.jx_nr_new,'9999')>0  order by  sbjx.jx_DATE DESC"
'sqljx=sqljx&" UNION ALL select "" as jx_id,sbgha.sb_id as sb_id,"" as jx_name,"" as jx_lb,sbgha.gh_yy as jx_gzxx,null as jx_gzxx_new,"" as jx_nr,"" as jx_nr_new,sbgha.gh_date as jx_date,"" as jx_enddate,"" as jx_fzren,sbgha.gh_ren as jx_ren,"" as jx_ylwt,sbgha.gh_bz as jx_bz,sbgha.gh_xh as gh_xh, sbgha.gh_xhupdate as gh_xhupdate,sb.sb_sscj,sb.sb_wh,sb.sb_dclass from sbgh as sbgha  left JOIN sb ON sbgha.sb_id = sb.sb_id"
  
  
'dwt.out sqljx



'	sqljx="SELECT sbjx.*,sbgh.gh_xh as gh_xh,sbgh.gh_xhupdate as gh_xhupdate from sbjx as sbjx"
'    sqljx= sqljx&" left join sbgh as sbgh on sbjx.jx_id=sbgh.jx_id where sbjx.sb_id="&sb_id&" and instr(sbjx.jx_nr_new,'9999')>0  order by  sbjx.jx_DATE DESC"
'    sqljx= sqljx&"   UNION ALL select null as jx_id,sbgha.sb_id as sb_id,null as jx_name,null as jx_lb,sbgha.gh_yy as jx_gzxx,null as jx_gzxx_new,null as jx_nr,null as jx_nr_new,sbgha.gh_date as jx_date,null as jx_enddate,null as jx_fzren,"
'    sqljx= sqljx&"	sbgha.gh_ren as jx_ren,null as jx_ylwt,sbgha.gh_bz as jx_bz,sbgha.gh_xh as gh_xh, sbgha.gh_xhupdate as gh_xhupdate from sbgh as sbgha where sbgha.sb_id="&sb_id

	
	
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
    message("û����ظ�����¼")
	else
		
		record=rsjx.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rsjx.PageSize = Cint(PgSz) 
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
		rsjx.absolutePage = page
		dim start
		start=PgSz*Page-PgSz+1
		rowCount = rsjx.PageSize
		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.out "     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>�豸λ��</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>�豸����</div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ʼʱ��</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>����ʱ��</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ʱ(Сʱ)</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>"& vbCrLf
		Dwt.out "    </tr>"& vbCrLf
		 do while not rsjx.eof and rowcount>0
			
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			
			if not isnull(rsjx("sb_sscj")) then
			   sscj=sscjh_d(rsjx("sb_sscj")) 
			 else
			  sscj=""
			 end if  
			
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscj&"</div></td>"& vbCrLf

			if not isnull( rsjx("jx_ylwt") )  then sbwh="<span style=""color:#ff0000"">��</span> " else sbwh=""	  '����������Ϊ��
			if not isnull(rsjx("sb_wh")) then 
			  sbwh=sbwh&"<a href=sb_ghjl.asp?sbid="&rsjx("sb_id")&"&sbclassid="&rsjx("sb_dclass")&">"&searchH(uCase(rsjx("sb_wh")),keys)&"</a>"
			else
			  sbwh=sbwh&"ɾ��"
			end if 
			
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" >"&sbwh&"</div></td>"& vbCrLf
			if rsjx("jx_lb")<>"" then 
    			jxlb=getjxlb(rsjx("jx_lb"))
            else
			    jxlb=""
			end if 
			
			if not isnull(rsjx("sb_wh")) then
     			sb_dclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&rsjx("sb_id"))(0))(0)
			else
			    	sb_dclassname=""
			end if 
			
						dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&sb_dclassname&"&nbsp;</td>"& vbCrLf

			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jxlb&"&nbsp;</td>"& vbCrLf
			
			
			
			
			
			jx_gzxx=""
			if not isnull( rsjx("jx_gzxx_new") ) then 
			  sbclassgz1=split(rsjx("jx_gzxx_new"),",")
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 then
							 '��ȡ���������� 
							 jxgzname=getjxgzxx(sbclassgz1(i))
							 'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +'��'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& jxgzname 
							else
							'��ȡ��������
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& "������"&rsjx("jx_gzxx") 
							end if 		  
							   		  
				   Next 
		    else
			    jx_gzxx="�����ݣ�"&rsjx("jx_gzxx") 
			end if  
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_gzxx&"&nbsp;</td>"& vbCrLf
			
			
			
			
			jx_nr=""
			if not isnull( rsjx("jx_nr_new") ) then 
			  sbclassgz1=split(rsjx("jx_nr_new"),",")
			   
			   
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 and sbclassgz1(i)<>9999 then     '0��������   99999�������
							 '��ȡ���������� 
								jxnrname=getjxnr(sbclassgz1(i))
								
										 'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +'��'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& jxnrname 
							else
							'��ȡ��������
								if sbclassgz1(i)=0 then 
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& "������"&rsjx("jx_nr")
								end if 
							'��ȡ��������
								if sbclassgz1(i)=9999 then 
										  if i<>0 then jx_nr=jx_nr & "<br>"
											'����Ҫ��� ��������Ϣ
										jx_nr=jx_nr& "������"
'										sqlgh="SELECT gh_xh,gh_xhupdate FROM sbgh  WHERE jx_id="&rsjx("jx_id")
'										set rsgh=server.createobject("adodb.recordset")
'										rsgh.open sqlgh,conn,1,1
'										if rsgh.eof and rsgh.bof then 
'												jx_nr=jx_nr&"δ�ҵ��������ͺ�����"
										'else
												jx_nr=jx_nr&"����ǰ�ͺ�<b>"&rsjx("gh_xh")&"</b>���������ͺ�<B>"&rsjx("gh_xhupdate")
'										end if   
		   
								end if 
							end if 			  
				   Next 
				   
				   'if jx_nr<> "" then 	       jx_nr=jx_nr&"<br>������"&rsjx("jx_nr") else  jx_nr="������"&rsjx("jx_nr") 
			else
			      ' (��ȡ�ɵ����ݻ���������
			      jx_nr="�����ݣ�"&"����ǰ�ͺ�<b>"&rsjx("gh_xh")&"</b>���������ͺ�<B>"&rsjx("gh_xhupdate")
			end if  
			

			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px;word-break:break-all;word-wrap:break-word"">"&jx_nr&"&nbsp;</td>"& vbCrLf
			jxdate=rsjx("jx_date")
			jxenddate=rsjx("jx_enddate")
			if year(jxdate)=year(now()) then jxdate=month(rsjx("jx_date"))&"-"&day(rsjx("jx_date"))&" "&hour(rsjx("jx_date"))&":"&minute(rsjx("jx_date"))
			if year(jxenddate)=year(now()) then jxenddate=month(rsjx("jx_enddate"))&"-"&day(rsjx("jx_enddate"))&" "&hour(rsjx("jx_enddate"))&":"&minute(rsjx("jx_enddate"))
			
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&jxdate&"</Div></td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&jxenddate&"</Div></td>"& vbCrLf
			

	a=rsjx("jx_date")
			b=rsjx("jx_enddate") 
			if not isnull(b) then 
			    ys=FormatNumber(DateDiff("n", a, b)/60,2,-1,0,0)
		     else
			    ys=""		
		     end if 		
			
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=""center"">"&ys&"</Div></td>"& vbCrLf
			
			
			
			jxfzren=""
			if not isnull( rsjx("jx_fzren") ) then 
			  sbclassgz1=split(rsjx("jx_fzren")," ")
			   
			   
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
				   jxfzren=jxfzren&"<a href=?keyword="&sbclassgz1(i)&">"&sbclassgz1(i)&"</a>&nbsp;&nbsp;"
				   next
			end if 	   
			
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&jxfzren&"&nbsp;</td>"& vbCrLf



			jxren=""
			if not isnull( rsjx("jx_ren") ) then 
			  sbclassgz1=split(rsjx("jx_ren")," ")
			   
			   
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
				   jxren=jxren&"<a href=?keyword="&sbclassgz1(i)&">"&sbclassgz1(i)&"</a>&nbsp;&nbsp;"
				   next
			end if 	   
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jxren&"&nbsp;</td>"& vbCrLf
			
			
			
			
			
			jx_ylwt=""
			if not isnull( rsjx("jx_ylwt") ) then 
			  sbclassgz1=split(rsjx("jx_ylwt"),",")
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							  
							 jxylwtname= getjxylwt(sbclassgz1(i))
							 ' conn.Execute("SELECT  sbjxylwt.sbjxylwt_name as sbjxylwt_name FROM sbjxylwt AS sbjxylwt left join sbjxylwt as sbjxylwtA on sbjxylwt.sbjxylwt_zclass=sbjxylwtA.sbjxylwt_id WHERE sbjxylwt.sbjxylwt_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_ylwt=jx_ylwt & "<br>" 
										  jx_ylwt=jx_ylwt& jxylwtname 
				   Next 
		    else
			    		jx_ylwt="��"
			end if  
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_ylwt&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjx("jx_bz")&"&nbsp;</td>"& vbCrLf
			Dwt.out" </tr>"			
			
			RowCount=RowCount-1
		rsjx.movenext
		loop
		Dwt.out"</table>"
		if sscjid<>"" or keys<>"" or request("action")="dcs" then 
				call showpage(page,url,total,record,PgSz)
        else
				call showpage1(page,url,total,record,PgSz)
		end if 
	end if
	rsjx.close
	set rsjx=nothing
	conn.close
	set conn=nothing







end sub	
sub ghcount 

url="sb_ghjl_left.asp?action=ghcount"
	sqlgh = "SELECT  sb_id,COUNT(sb_id) as sbghnumb FROM sbgh  GROUP BY sb_id order by COUNT(sb_id) desc"
	set rsgh=server.createobject("adodb.recordset")
	rsgh.open sqlgh,conn,1,1
	if rsgh.eof and rsgh.bof then 
		message("û�м��޼�¼")
	else
		record=rsgh.recordcount
		
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		
		rsgh.PageSize = Cint(PgSz) 
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
		rsgh.absolutePage = page
		dim start
		start=PgSz*Page-PgSz+1
		rowCount = rsgh.PageSize
	
	
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td class='x-td'><DIV class='x-grid-hd-text'>���</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>����</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>�豸λ��</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>�豸����</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>���޴���</div></td>"& vbCrLf
		dwt.out "    </tr>"& vbCrLf
	do while not rsgh.eof and rowcount>0
		'dwt.out rsgh("sb_id")&" "
'		sql = "SELECT sbgh.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbgh where sbgh.sb_id="&rsgh("sb_id")&" INNER JOIN sb ON sbgh.sb_id = sb.sb_id "
'		set rs=server.createobject("adodb.recordset")
'		rsgh.open sql,conn,1,1
'		if rs.eof and rs.bof then 
'			message("û�м��޼�¼")
'		else
		
		Set Rs =conn.Execute("SELECT sb_wh,sb_sscj FROM sb WHERE sb_id="&rsgh("sb_id"))
		if rs.eof and rs.bof then 
		else
			sb_sscj=rs("sb_sscj")
			'dwt.out sb_sscj
			ghnumb=conn.Execute("SELECT count(sb_id) FROM sbgh WHERE sb_id="&rsgh("sb_id"))(0)
			sb_dclass=conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&rsgh("sb_id"))(0)
			sb_dclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&rsgh("sb_id"))(0))(0)
		    
			sb_wh=rs("sb_wh")
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_d(sb_sscj)&"</div></td>"& vbCrLf
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><a href=sb_ghjl.asp?sbid="&rsgh("sb_id")&"&sbclassid="&sb_dclass&">"&searchH(uCase(sb_wh),keys)&"</div></td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&sb_dclassname&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&ghnumb&"&nbsp;</td>"& vbCrLf
			 dwt.out "</tr>"	
			RowCount=RowCount-1
		end if 


	rsgh.movenext
	loop
	dwt.out"</table>"
	call showpage(page,url,total,record,PgSz)
	end if 



end sub


dwt.out "</body></html>"




sub search()
	dim sqlcj,rscj
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='sb_ghjl_left.asp'>" & vbCrLf
	
    dwt.out "&nbsp;&nbsp;<input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 dwt.out "value='"&request("keyword")&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='����������λ��'"
	 dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	dwt.out "  <input type='submit' name='Submit'  value='����'>&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>��������ת����</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='sb_ghjl_left.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
			dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "</select>	" & vbCrLf
	dwt.out "<a href=sb_ghjl_left.asp?action=ghcount>��������������</a>  <a href=?action=dcs>ֻ��ʾDCS���</a>"
	dwt.out "</form></div></div>" & vbCrLf
end sub

Call CloseConn
%>