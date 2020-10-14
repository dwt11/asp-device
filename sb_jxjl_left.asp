<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/sb_function.asp"-->
<%dim url
dim sqlbody,rsbody,ii,sbclassid,ylbid,sqljx,rsjx
dim record,pgsz,total,page,rowCount,xh,sscj
dim sb_wh,sql,rs


url=geturl
keys=ReplaceBadChar(trim(request("keyword")) )
sscjid=trim(request("sscj")) 


jx_date=trim(request("jx_date")) 
jx_enddate=trim(request("jx_enddate")) 
dwt.out"<html>"& vbCrLf
dwt.out"<head>" & vbCrLf
dwt.out"<title>检修记录汇总</title>"& vbCrLf
dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out"<script language='javascript' type='text/javascript' src='js/My97DatePicker/WdatePicker.js'></script>"
dwt.out"</head>"& vbCrLf
dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf


	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>检修记录汇总&nbsp;&nbsp;&nbsp;&nbsp;"& vbCrLf
   	dwt.out "     </span></div>"& vbCrLf
   call search()


if request("action")="" or request("action")="dcs" then call main
if request("action")="jxcount" then call jxcount
sub main()
  
  
  
  
  
  
if request("action")="dcs" then 

    sqljx = "SELECT sbjx.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id where sb.sb_dclass=121 or sb.sb_dclass=105 or sb.sb_dclass=144 or sb.sb_dclass=123 or sb.sb_dclass=124 or sb.sb_dclass=125 or sb.sb_dclass=127 or sb.sb_dclass=128 or sb.sb_dclass=129 or sb.sb_dclass=133 or sb.sb_dclass=134 or sb.sb_dclass=126 or sb.sb_dclass=84 or sb.sb_dclass=88 or sb.sb_dclass=91 or sb.sb_dclass=94 or sb.sb_dclass=95 or sb.sb_dclass=96 or sb.sb_dclass=97 or sb.sb_dclass=98 or sb.sb_dclass=99 "
 else
  
  sqljx = "SELECT sbjx.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id "
 end if 
  sqlwhere=" where 1=1 "
  if sscjid<>"" then sqlwhere=sqlwhere&" and (((sb.sb_sscj)="&sscjid&")) "
  if keys<>"" then sqlwhere=sqlwhere&" and ( sb.sb_wh like '%" &keys& "%' or sbjx.jx_ren like '%" &keys& "%'  or sbjx.jx_fzren like '%" &keys& "%' )"
if jx_date<>"" or jx_enddate<>"" then sqlwhere=sqlwhere&" and  (jx_date between #"&jx_date&"# and #"&jx_enddate&"# ) "



  sqljx=sqljx& sqlwhere &" order by  jx_date DESC"	
  
  
'  dwt.out sqljx
	'sqljx="SELECT * from sbjx where sb_id="&sb_id&" order by  jx_DATE DESC"
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
		message("没有检修记录")
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
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>设备位号</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>设备类型</div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td' ><Div class='x-grid-hd-text'>检修类别</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>故障现象</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>检修内容</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>开始时间</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>结束时间</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>用时(小时)</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>负责人</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>检修人</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>遗留问题</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>备注</Div></td>"& vbCrLf
		Dwt.out "    </tr>"& vbCrLf
		 do while not rsjx.eof and rowcount>0
		sqlgh2="SELECT sbclass_name from sbclass where sbclass_id="&rsjx("sb_dclass")
		set rsgh2=server.createobject("adodb.recordset")
		rsgh2.open sqlgh2,conn,1,1
		if rsgh2.eof and rsgh2.bof then 
		   sb_classname="在设备管理中此位号已经删除"
		else
		   sb_classname=rsgh2("sbclass_name")
		end if 
		rsgh2.close
		
					xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
			
			if not isnull(rsjx("sb_sscj")) then
			   sscj=sscjh_d(rsjx("sb_sscj")) 
			 else
			  sscj=""
			 end if  
			
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscj&"</div></td>"& vbCrLf
			
			if not isnull( rsjx("jx_ylwt") )  then sbwh="<span style=""color:#ff0000"">★</span> " else sbwh=""	  '有遗留问题为红
			if not isnull(rsjx("sb_wh")) then 
			  sbwh=sbwh&"<a href=sb_jxjl.asp?sbid="&rsjx("sb_id")&"&sbclassid="&rsjx("sb_dclass")&">"&searchH(uCase(rsjx("sb_wh")),keys)&"</a>"
			else
			  sbwh=sbwh&"删除"
			end if 
			
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" >"&sbwh&"</div></td>"& vbCrLf

			
			
			
			
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&sb_classname&"&nbsp;</td>"& vbCrLf
				jxlb=""
			if rsjx("jx_lb")<>"" then 
    			jxlb=jxlb&getjxlb(rsjx("jx_lb"))
            else
			    jxlb=jxlb&""
			end if 
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jxlb&"&nbsp;</td>"& vbCrLf
			
			
			
			
			
			jx_gzxx=""
			if not isnull( rsjx("jx_gzxx_new") ) then 
			  sbclassgz1=split(rsjx("jx_gzxx_new"),",")
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 then
							 '读取正常的数据 
							 jxgzname=getjxgzxx(sbclassgz1(i))
							 'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +'：'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& jxgzname 
							else
							'读取其他数据
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& "其他："&rsjx("jx_gzxx") 
							end if 		  
							   		  
				   Next 
'				   if jx_gzxx<>"" then  
'				         jx_gzxx=jx_gzxx&"<br>其他："&rsjx("jx_gzxx") 
'					else 
'					     jx_gzxx="其他："&rsjx("jx_gzxx")
'					end if 	 
		    else
			    jx_gzxx="旧数据："&rsjx("jx_gzxx") 
			end if  
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_gzxx&"&nbsp;</td>"& vbCrLf
			
			
			
			
			jx_nr=""
			if not isnull( rsjx("jx_nr_new") ) then 
			  sbclassgz1=split(rsjx("jx_nr_new"),",")
			   
			   
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 and sbclassgz1(i)<>9999 then     '0代表其他   99999代表更换
							 '读取正常的数据 
								jxnrname=getjxnr(sbclassgz1(i))
								
										 'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +'：'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& jxnrname 
							else
							'读取其他数据
								if sbclassgz1(i)=0 then 
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& "其他："&rsjx("jx_nr")
								end if 
							'读取更换数据
								if sbclassgz1(i)=9999 then 
										  if i<>0 then jx_nr=jx_nr & "<br>"
											'这里要检测 更换的信息
										jx_nr=jx_nr& "更换："
										sqlgh="SELECT gh_xh,gh_xhupdate FROM sbgh  WHERE jx_id="&rsjx("jx_id")
										set rsgh=server.createobject("adodb.recordset")
										rsgh.open sqlgh,conn,1,1
										if rsgh.eof and rsgh.bof then 
												jx_nr=jx_nr&"未找到更换的型号数据"
										else
'											 if jx_nr= "" then 
'												jx_nr="更换：更换前型号<b>"&ghxh&"</b>，更换后型号<B>"&ghxhupdate
'											 else
												jx_nr=jx_nr&"更换前型号<b>"&rsgh("gh_xh")&"</b>，更换后型号<B>"&rsgh("gh_xhupdate")
'											 end if 
										end if   
		   
								end if 
							end if 			  
				   Next 
				   
				   'if jx_nr<> "" then 	       jx_nr=jx_nr&"<br>其他："&rsjx("jx_nr") else  jx_nr="其他："&rsjx("jx_nr") 
			else
			      ' (读取旧的数据或其他数据
			      jx_nr="旧数据："&rsjx("jx_nr")
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
			    		jx_ylwt="无"
			end if  
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_ylwt&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjx("jx_bz")&"&nbsp;</td>"& vbCrLf
			Dwt.out" </tr>"			
			
			RowCount=RowCount-1
		rsjx.movenext
		loop
		Dwt.out"</table>"
		if sscjid<>"" or keys<>"" OR request("action")="dcs" then 
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


sub jxcount 

url="sb_jxjl_left.asp?action=jxcount"
	sqljx = "SELECT  sb_id,COUNT(sb_id) as sbjxnumb FROM sbjx  GROUP BY sb_id order by COUNT(sb_id) desc"
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
		message("没有检修记录")
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
	
	
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>车间</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>设备位号</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>设备类型</div></td>"& vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>检修次数</div></td>"& vbCrLf
		dwt.out "    </tr>"& vbCrLf
	do while not rsjx.eof and rowcount>0
		'dwt.out rsjx("sb_id")&" "
'		sql = "SELECT sbjx.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbjx where sbjx.sb_id="&rsjx("sb_id")&" INNER JOIN sb ON sbjx.sb_id = sb.sb_id "
'		set rs=server.createobject("adodb.recordset")
'		rsjx.open sql,conn,1,1
'		if rs.eof and rs.bof then 
'			message("没有检修记录")
'		else
		
		Set Rs =conn.Execute("SELECT sb_wh,sb_sscj FROM sb WHERE sb_id="&rsjx("sb_id"))
		if rs.eof and rs.bof then 
		else
			sb_sscj=rs("sb_sscj")
			'dwt.out sb_sscj
			jxnumb=conn.Execute("SELECT count(sb_id) FROM sbjx WHERE sb_id="&rsjx("sb_id"))(0)
			sb_dclass=conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&rsjx("sb_id"))(0)
			sb_dclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&rsjx("sb_id"))(0))(0)
		    
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
			dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"" ><a href=sb_jxjl.asp?sbid="&rsjx("sb_id")&"&sbclassid="&sb_dclass&">"&searchH(uCase(sb_wh),keys)&"</div></td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&sb_dclassname&"&nbsp;</td>"& vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jxnumb&"&nbsp;</td>"& vbCrLf
			 dwt.out "</tr>"	
			RowCount=RowCount-1
		end if 


	rsjx.movenext
	loop
	dwt.out"</table>"
	call showpage(page,url,total,record,PgSz)
	end if 



end sub





dwt.out "</body></html>"




sub search()
	dim sqlcj,rscj
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='sb_jxjl_left.asp'>" & vbCrLf
	
    dwt.out "&nbsp;&nbsp;<input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 dwt.out "value='"&request("keyword")&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='输入搜索的关键字'"
	 dwt.out " onblur=""if(this.value==''){this.value='输入搜索的关键字'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
Dwt.out"<input name='jx_date'  type='text'  id='jx_date'  class='Wdate' onFocus=""var jx_enddate=$dp.$('jx_enddate');WdatePicker({onpicked:function(){pickedFunc();jx_enddate.focus();},dateFmt:'yyyy/MM/dd ',maxDate:'#F{$dp.$D(\'jx_enddate\')}'})""   readOnly  value='"&jx_date&"'>"
   dwt.out " 至 "
   
   Dwt.out"<input name='jx_enddate' type='text'  id='jx_enddate'  class='Wdate'   onFocus=""WdatePicker({dateFmt:'yyyy/MM/dd ',minDate:'#F{$dp.$D(\'jx_date\')}',onpicked:pickedFunc})""   readOnly  value='"&jx_enddate&"'>"
 %>
   <script language="JavaScript">
    // 计算两个日期的间隔天数  
     //   document.all.dateChangDu.value = iDays;
	function pickedFunc(){
		  Date.prototype.dateDiff = function(interval,objDate){    
		//若参数不足或 objDate 不是日期物件则回传 undefined    
		if(arguments.length<2||objDate.constructor!=Date) return undefined;    
		  }
	
	}

</script>  
 <%  
	dwt.out "  <input type='submit' name='Submit'  value='搜索'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='sb_jxjl_left.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
			dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "</select>	" & vbCrLf
	
	dwt.out "<a href=sb_jxjl_left.asp?action=jxcount>按检修次数排列</a> <a href=sb_jxjl_left.asp?action=dcs>只显示DCS相关</a>"
	
	dwt.out "</form></div></div>" & vbCrLf
end sub


Call CloseConn
%>