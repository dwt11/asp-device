<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/sb_function.asp"-->
<%dim url
dim sqlbody,rsbody,ii,sb_classid,sbid,sqlgh,rsgh
dim record,pgsz,total,page,rowCount,xh,sb_sscj
dim sb_wh,sql,rs

sb_id=Trim(Request("sbid"))
sbclass_id=Trim(Request("sbclassid"))
url="sb_ghjl.asp?sbid="&sb_id&"&sbclassid="&sbclass_id
'读取分类，以用于标题
if sbclass_id="" or sb_id="" then Dwt.out"<Script Language=Javascript>history.back()</Script>"
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sbclass_id)(0)
'if sb_id<>"" then 
 sb_wh=conn.Execute("SELECT sb_wh FROM sb WHERE  sb_id="&sb_id)(0)
 sb_sscj=conn.Execute("SELECT sb_sscj FROM sb WHERE  sb_id="&sb_id)(0)
'end if 
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>技术档案管理页</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&sb_classname&"  "&sb_wh&" 更换记录</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf

action=request("action")
select case action
  case ""
      call main
end select	  	 


	
sub main()
    
	
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<a href='sb.asp?sbclassid="&sbclass_id&"&keyword="&sb_wh&"'>点击查看 "&sb_wh&" 的详细信息</a> "
	
	'此句是不兼容旧更换数据 的
	'sqljx="SELECT * from sbjx where (sb_id="&sb_id&" and instr(jx_nr_new,'9999')>0)   order by  jx_DATE DESC"
	
	
	'兼容旧的更换数据表
	'sqljx="SELECT * from sbjx where sb_id="&sb_id&" and instr(jx_nr_new,'9999')>0  order by  jx_DATE DESC"
	'sqljx= sqljx&" union all select "
	'sqljx= sqljx&" '' as jx_id,sb_id,'' as jx_name,'' as jx_lb,gh_yy as jx_gzxx,'' as jx_gzxx_new,'' as jx_nr,'' as jx_nr_new,gh_date as jx_date,'' as jx_enddate,'' as jx_fzren,gh_ren as jx_ren,'' as jx_ylwt,"
	'sqljx= sqljx&" gh_bz as jx_bz from sbgh where sb_id="&sb_id

	sqljx="SELECT sbjx.*,sbgh.gh_xh as gh_xh,sbgh.gh_xhupdate as gh_xhupdate from sbjx as sbjx"
    sqljx= sqljx&" left join sbgh as sbgh on sbjx.jx_id=sbgh.jx_id where sbjx.sb_id="&sb_id&" and instr(sbjx.jx_nr_new,'9999')>0  order by  sbjx.jx_DATE DESC"
   ' sqljx= sqljx&"   UNION ALL select null as jx_id,sbgha.sb_id as sb_id,null as jx_name,null as jx_lb,sbgha.gh_yy as jx_gzxx,null as jx_gzxx_new,null as jx_nr,null as jx_nr_new,sbgha.gh_date as jx_date,null as jx_enddate,null as jx_fzren,"
   ' sqljx= sqljx&"	sbgha.gh_ren as jx_ren,null as jx_ylwt,sbgha.gh_bz as jx_bz,sbgha.gh_xh as gh_xh, sbgha.gh_xhupdate as gh_xhupdate from sbgh as sbgha where sbgha.sb_id="&sb_id &"   order by jx_date  desc "

	
	
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
		
		Dwt.out"<input name='Cancel' type='button' id='Cancel' value=' 返  回 ' onClick="";history.back()"" style='cursor:hand;'>"
		Dwt.out "</Div></Div>"
		message("未添加  "&sb_wh&" 检修记录")
	else
		
		Dwt.out"<input name='Cancel' type='button' id='Cancel' value=' 返  回 ' onClick="";history.back()"" style='cursor:hand;'>"
		Dwt.out "</Div></Div>"
		
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
			
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
			
			
			jxlb=""
			if not isnull( rsjx("jx_ylwt") )  then jxlb="<span style=""color:#ff0000"">★</span> "  	  '有遗留问题为红

			
			
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
'										sqlgh="SELECT gh_xh,gh_xhupdate FROM sbgh  WHERE jx_id="&rsjx("jx_id")
'										set rsgh=server.createobject("adodb.recordset")
'										rsgh.open sqlgh,conn,1,1
'										if rsgh.eof and rsgh.bof then 
'												jx_nr=jx_nr&"未找到更换的型号数据"
										'else
												jx_nr=jx_nr&"更换前型号<b>"&rsjx("gh_xh")&"</b>，更换后型号<B>"&rsjx("gh_xhupdate")
'										end if   
		   
								end if 
							end if 			  
				   Next 
				   
				   'if jx_nr<> "" then 	       jx_nr=jx_nr&"<br>其他："&rsjx("jx_nr") else  jx_nr="其他："&rsjx("jx_nr") 
			else
			      ' (读取旧的数据或其他数据
			      jx_nr="旧数据："&"更换前型号<b>"&rsjx("gh_xh")&"</b>，更换后型号<B>"&rsjx("gh_xhupdate")
			end if  
			

			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px;word-break:break-all;word-wrap:break-word"">"&jx_nr&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsjx("jx_date")&"</Div></td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsjx("jx_enddate")&"</Div></td>"& vbCrLf
			
			a=rsjx("jx_date")
			b=rsjx("jx_enddate") 
			if not isnull(b) then 
			    ys=FormatNumber(DateDiff("n", a, b)/60,2,-1,0,0)
		     else
			    ys=""		
		     end if 		
			
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=""center"">"&ys&"</Div></td>"& vbCrLf
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rsjx("jx_fzren")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjx("jx_ren")&"&nbsp;</td>"& vbCrLf
			
			
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
		call showpage(page,url,total,record,PgSz)
	end if
	rsjx.close
	set rsjx=nothing
	conn.close
	set conn=nothing
		
		message "(添加更换记录)已经合并入(添加检修记录)页面,请在添加检修记录的时候选择(检修内容)为更换.<br>此处仅做为当前设备的更换记录汇总显示."
end sub

Call CloseConn
%>