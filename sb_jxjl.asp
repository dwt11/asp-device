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
dim sqlbody,rsbody,ii,sbclassid,ylbid,sqljx,rsjx
dim record,pgsz,total,page,rowCount,xh,sscj
dim sb_wh,sql,rs


sb_id=Trim(Request("sbid"))
sbclass_id=Trim(Request("sbclassid"))
url="sb_jxjl.asp?sbid="&sb_id&"&sbclassid="&sbclass_id
'读取分类，以用于标题
if sbclass_id="" or sb_id="" then Dwt.out"<Script Language=Javascript>history.back()</Script>"
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sbclass_id)(0)
'if sb_id<>"" then 
 sb_wh=conn.Execute("SELECT sb_wh FROM sb WHERE  sb_id="&sb_id)(0)
 sb_sscj=conn.Execute("SELECT sb_sscj FROM sb WHERE  sb_id="&sb_id)(0)
'end if 
Dwt.out"<html>"& vbCrLf
Dwt.out"<head>" & vbCrLf
Dwt.out"<title> 技术档案管理页</title>"& vbCrLf
Dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function CheckAdd(){" & vbCrLf


%>
        var checkName = document.getElementsByName ("jx_gzxx_new");	//根据组件名获取组建对象
		var ischecked
		//循环checkbox，判断是否包含选中项
		for (i = 0; i < checkName.length; i ++) {
			if (checkName[i].checked) {	//如果有选中项，则返回true
				ischecked= true;
                
                if(checkName[i].value==0){
                  if(document.getElementById("jx_gzxx").value==''){
                         alert('故障现象不能为空！');
                      document.getElementById("jx_gzxx").focus();
                          return false;
                    }
                }
			}
		}
		if(!ischecked)
		 {alert('未选择故障现象！');
		  return false;	//循环结束，无选中项，返回FALSE
		 }

         checkName = document.getElementsByName ("jx_nr_new");	//根据组件名获取组建对象
		 ischecked=false
		//循环checkbox，判断是否包含选中项
		for (i = 0; i < checkName.length; i ++) {
			if (checkName[i].checked) {	//如果有选中项，则返回true
				ischecked= true;
                if(checkName[i].value==0){
                  if(document.getElementById("jx_nr").value==''){
                         alert('检修内容不能为空！');
                      document.getElementById("jx_nr").focus();
                          return false;
                    }
                }
                if(checkName[i].value==9999){
                  if(document.getElementById("gh_xh").value==''){
                         alert('更换前型号不能为空！');
                      document.getElementById("gh_xh").focus();
                          return false;
                    }
                  if(document.getElementById("gh_xhupdate").value==''){
                         alert('更换后型号不能为空！');
                      document.getElementById("gh_xhupdate").focus();
                          return false;
                    }
                }
                
                
                
			}
		}
		if(!ischecked)
		 {alert('未选择检修内容！');
		  return false;	//循环结束，无选中项，返回FALSE
		 }

<%




Dwt.out "  if(document.formadd.jx_fzren.value==''){" & vbCrLf
Dwt.out "      alert('检修负责人不能为空！');" & vbCrLf
Dwt.out "  document.formadd.jx_fzren.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "  if(document.formadd.jx_ren.value==''){" & vbCrLf
Dwt.out "      alert('检修人不能为空！');" & vbCrLf
Dwt.out "  document.formadd.jx_ren.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "  if(document.formadd.jx_date.value==''){" & vbCrLf
Dwt.out "      alert('检修时间不能为空！');" & vbCrLf
Dwt.out "  document.formadd.jx_date.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

'Dwt.out "execScript('t=IsDate(document.formadd.jx_date.value)','VBScript');" & vbCrLf
'Dwt.out "if(!t){" & vbCrLf
'Dwt.out"   alert('日期格式不正确，应用yyyy-mm-dd');" & vbCrLf
'Dwt.out "  document.formadd.jx_date.focus();" & vbCrLf
'Dwt.out "return false;" & vbCrLf
'Dwt.out "    }" & vbCrLf


Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out"<script language='javascript' type='text/javascript' src='js/My97DatePicker/WdatePicker.js'></script>"
Dwt.out"</head>"& vbCrLf
action=request("action")

Dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'"
'如果是编辑 则JS加载 用时计算 ,其他和更换的框判断显示
if action="edit" then dwt.out " onload='pickedFunc();clickgzxxqt();clickjxnrqt();clickjxnrgh()' "
'如果是增加 则JS加载 用时计算
if action="add" then dwt.out " onload='pickedFunc();' "
dwt.out ">"& vbCrLf

	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&sb_classname&"  "&sb_wh&" 检修记录</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf


select case action
  case "add"
      call add'添加设备分类选择
  case "saveadd"
      call saveadd'添加设备分类选择
  case "edit"
      call edit
  case "saveedit"'编辑子分类
      call saveedit'编辑保存子分类
  case "del"
      call del     '删除子分类信息
  case ""
      call main
end select	  	 




sub main()
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<a href='sb.asp?sbclassid="&sbclass_id&"&keyword="&sb_wh&"'>点击查看 "&sb_wh&" 的详细信息</a> "

	sqljx="SELECT * from sbjx where sb_id="&sb_id&" order by  jx_DATE DESC"
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
		if session("levelclass")=sb_sscj or session("levelclass")=0 then
			Dwt.out"<input type='button' name='Submit'  onclick=""window.location.href='sb_jxjl.asp?action=add&sbid="&sb_id&"&sbclassid="&sbclass_id&"'""value='添加检修记录'>"
		end if 	
		Dwt.out"<input name='Cancel' type='button' id='Cancel' value=' 返  回 ' onClick="";history.back()"" style='cursor:hand;'>"
		Dwt.out "</Div></Div>"
		message("未添加  "&sb_wh&" 检修记录")
	else
		if session("levelclass")=sb_sscj or session("levelclass")=0 then
			Dwt.out"<input type='button' name='Submit'  onclick=""window.location.href='sb_jxjl.asp?action=add&sbid="&sb_id&"&sbclassid="&sbclass_id&"'""value='添加检修记录'>"
		end if 	
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
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>选项</Div></td>"& vbCrLf
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
			Dwt.out" <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
			call jxxuanxiang(rsjx("jx_id"),sb_id,sb_sscj)
			Dwt.out"</Div></td></tr>"			
			
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
end sub


sub add()
   '新增位号检修记录
   Dwt.out"<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'  >"& vbCrLf
   Dwt.out"<form method='post' action='sb_jxjl.asp' name='formadd' onsubmit='javascript:return CheckAdd();'>"& vbCrLf
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.out"<Div align='center'><strong>新增   "&sb_wh&" 检修记录</strong></Div></td>    </tr>"& vbCrLf
  
  
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修类别： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
  dim ischecked  '用于判断检修类别第一项 是否默认选中
  ischecked=false
   
    dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxlb where sbjxlb_zclass=0 order by  sbjxlb_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">暂无内容</p>" 
  else
	  do while not rsbody.eof 
						'二级
				sqlz="SELECT * from sbjxlb where sbjxlb_zclass="&rsbody("sbjxlb_id")&" order by  sbjxlb_orderby aSC"& vbCrLf
				set rsz=server.createobject("adodb.recordset")
				rsz.open sqlz,conn,1,1
				if rsz.eof and rsz.bof then 
					dwt.out"<input type='radio' name='jx_lb' value='"&rsbody("sbjxlb_id")&"'" 
							if not ischecked then 
							   ischecked=true
							   dwt.out "checked"
							end if    
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name") & "<br>"
				else
					do while not rsz.eof
					
						dwt.out"<input type='radio' name='jx_lb' value='"&rsz("sbjxlb_id")&"'" 
							if not ischecked then 
							   ischecked=true
							   dwt.out "checked"
							end if    
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name")&":"&rsz("sbjxlb_name") & "<br>"
					rsz.movenext
					loop
				end if 	
				rsz.close
				set rsz=nothing
			
		rsbody.movenext
		loop
  end if 
  rsbody.close
  set rsbody=nothing
   
   
   
		
  

   dwt.out "</td></tr>"& vbCrLf

  
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>故障现象： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassgz,jxgzname
   sbclassgz=conn.Execute("SELECT sb_jxgzxx_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   if not isnull( sbclassgz ) then 
	  sbclassgz1=split(sbclassgz,",")
	 For i = LBound(sbclassgz1) To UBound(sbclassgz1)
              	
				jxgzname=getjxgzxx(sbclassgz1(i))
				'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +':'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_gzxx_new' value='"&sbclassgz1(i)&"'>"	
				dwt.out i+1& "-"&jxgzname & "<br>"
				
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_gzxx_new' value='0' onclick='clickgzxxqt()'>其他<span id=jxgzxxspan></span>"	
			
   dwt.out "</td></tr>"& vbCrLf
   



   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修内容： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassjx,jxnrname
   sbclassjx=conn.Execute("SELECT sb_jxnr_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   
   if not isnull( sbclassjx ) then 
	  sbclassjx1=split(sbclassjx,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
                jxnrname=getjxnr(sbclassjx1(i))
				'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +':'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassjx1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_nr_new' value='"&sbclassjx1(i)&"'>"	
				dwt.out i+1& "-"&jxnrname & "<br>"
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_nr_new' value='0'  id='jxnrqt' onclick='clickjxnrqt()' >其他<span id=jxnrspan></span>"	
   dwt.out"<br><input type='checkbox' name='jx_nr_new' value='9999' onclick='clickjxnrgh()'>更换<span id='jxnrgh'></span>"	
   dwt.out "</td></tr>"& vbCrLf
   %>
      <script language="JavaScript">
		
		//检测故障显示的其他是否选中
		function clickgzxxqt(){
		  var checkName = document.getElementsByName ("jx_gzxx_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==0)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			jxgzxxspan.innerHTML="：<input name='jx_gzxx' id='jx_gzxx' type='text'>"
			}else{
			jxgzxxspan.innerHTML=""
		  }
		}
		//检测检修内容的其他是否选中
		function clickjxnrqt(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==0)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			jxnrspan.innerHTML="：<input name='jx_nr' id='jx_nr' type='text'>"
			}else{
			jxnrspan.innerHTML=""
		  }
		}
		//检测检修内容的更换是否选中
		function clickjxnrgh(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==9999)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			  <%
			  ghqxh=conn.Execute("SELECT sb_ggxh FROM sb  WHERE sb_id="&sb_id)(0)
			  
			  %>
			jxnrgh.innerHTML="：更换前型号 <input name='gh_xh' id='gh_xh'  type='text' value='<%=ghqxh%>'>&nbsp;更换后型号<input name='gh_xhupdate'  id='gh_xhupdate'  type='text'>"
			}else{
			jxnrgh.innerHTML=""
		  }
		}
      </script>
   <%
   
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修时间： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   Dwt.out"<input name='jx_date'  type='text'  id='jx_date'  class='Wdate' onFocus=""var jx_enddate=$dp.$('jx_enddate');WdatePicker({onpicked:function(){pickedFunc();jx_enddate.focus();},dateFmt:'yyyy/MM/dd HH:mm',maxDate:'#F{$dp.$D(\'jx_enddate\')}'})""   readOnly  value='"&now()&"'>"
   dwt.out " 至 "
   
   Dwt.out"<input name='jx_enddate' type='text'  id='jx_enddate'  class='Wdate'   onFocus=""WdatePicker({dateFmt:'yyyy/MM/dd HH:mm',minDate:'#F{$dp.$D(\'jx_date\')}',onpicked:pickedFunc})""   readOnly  value='"&now()&"'>"
   
   dwt.out "&nbsp;&nbsp;<span id='jxys'></span>  "
   
   Dwt.out"</td></tr>"& vbCrLf
   %>
   <script language="JavaScript">
    // 计算两个日期的间隔天数  
     //   document.all.dateChangDu.value = iDays;
	function pickedFunc(){
		  Date.prototype.dateDiff = function(interval,objDate){    
		//若参数不足或 objDate 不是日期物件则回传 undefined    
		if(arguments.length<2||objDate.constructor!=Date) return undefined;    
		switch (interval) {      
		//计算秒差    
		 // case "s":return parseInt((objDate-this)/1000);      
		  //计算分差    
			case "n":return parseInt(Math.round(((objDate-this)/60000)*100)/100);      
			//计算时差    
			  case "h":return Math.round(((objDate-this)/3600000)*100)/100;      
			  //计算日差      
			 // case "d":return parseInt((objDate-this)/86400000);      
			  //计算月差      
			 // case "m":return (objDate.getMonth()+1)+((objDate.getFullYear()-this.getFullYear())*12)-(this.getMonth()+1);      
			  //计算年差      
			 // case "y":return objDate.getFullYear()-this.getFullYear();      
			
			  //输入有误      
			  default:return undefined;    
			}
		 }
		//document.all.dateChangDu.value = document.all.jx_date.value;
			  var sDT = new Date(document.all.jx_date.value);
			  var eDT = new Date(document.all.jx_enddate.value);
			  jxys.innerHTML=("用时："+ sDT.dateDiff("h",eDT)+"小时");
	
	}

</script>  
   
   <%
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修负责人： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_fzren' type='text' ></td></tr>"& vbCrLf
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检 修 人： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_ren' type='text'><br>两个字的姓名,两字中间请勿添加空格或其他字符 <br>多个检修人,每个人的姓名中间请用空格区分,请勿使用其他字符</td></tr>"& vbCrLf
   
      Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>遗留问题： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   		
					sqlz="SELECT * from sbjxylwt order by  sbjxylwt_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
						do while not rsz.eof
							dwt.out"<input type='checkbox' name='jx_ylwt' value='"&rsz("sbjxylwt_id")&"'" 
											Dwt.out ">"	
											dwt.out rsz("sbjxylwt_name") & "<br>"
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing 
  

   dwt.out "<b>上述都不选择表示无遗留问题</b>"& vbCrLf
   dwt.out "<br><b>如果选择上述项目则表示有遗留问题,设备""完好""状态会更新为不完好</b></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备    注： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_bz' type='text'></td></tr>"& vbCrLf
   
   
   Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveadd'> <input name='sbid' type='hidden'  value='"&Trim(Request("sbid"))&"'> <input name='sbclassid' type='hidden'  value='"&Trim(Request("sbclassid"))&"'> <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='sb_jxjl.asp?sbid="&Trim(Request("sbid"))&"&sbclassid="&Trim(Request("sbclassid"))&"';"" style='cursor:hand;'></td>  </tr>"

   Dwt.out"</form></table>"& vbCrLf
end sub	

sub saveadd()    
	'保存新增检修记录
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sbjx" 
	rsadd.open sqladd,conn,1,3
	rsadd.addnew
	rsadd("jx_lb")=ReplaceBadChar(Trim(Request("jx_lb")))
	rsadd("jx_gzxx")=ReplaceBadChar(Trim(Request("jx_gzxx")))
	rsadd("jx_nr")=ReplaceBadChar(Trim(request("jx_nr")))
	rsadd("jx_gzxx_new")=ReplaceBadChar(Trim(Request("jx_gzxx_new")))
	rsadd("jx_nr_new")=ReplaceBadChar(Trim(request("jx_nr_new")))
	rsadd("jx_date")=Trim(request("jx_date"))
	rsadd("jx_enddate")=Trim(request("jx_enddate"))
	rsadd("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsadd("jx_ren")=ReplaceBadChar(Trim(request("jx_ren")))
	rsadd("jx_ylwt")=ReplaceBadChar(Trim(request("jx_ylwt")))
	rsadd("jx_bz")=ReplaceBadChar(Trim(request("jx_bz")))
	rsadd("sb_id")=ReplaceBadChar(Trim(request("sbid")))
	rsadd.update
	jxid= rsadd("jx_id")
	rsadd.close
	
	
	ghxh=ReplaceBadChar(Trim(request("gh_xh")))
	ghxhupdate=ReplaceBadChar(Trim(request("gh_update")))
   dim isgh '是否选中更换
  isgh =false
   if not isnull( ReplaceBadChar(Trim(request("jx_nr_new"))) ) then 
	  sbclassjx1=split(ReplaceBadChar(Trim(request("jx_nr_new"))),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=9999 then isgh=true
	 Next 
	 
	end if  

	'如果添加了更换数据
	if isgh and jxid<>"" then 
	  '保存新增更换记录
      set rsaddgh=server.createobject("adodb.recordset")
      sqladdgh="select * from sbgh" 
      rsaddgh.open sqladdgh,conn,1,3
      rsaddgh.addnew
      rsaddgh("jx_id")=jxid
      rsaddgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
      rsaddgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
     rsaddgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
     rsaddgh.update
      rsaddgh.close
      set rsaddgh=nothing
	  
	   '更改设备档案中相应位号设备的规格型号2008-9-19
	  set rsadd1=server.createobject("adodb.recordset")
          sqladd1="select * from sb where sb_id="&Trim(request("sbid"))
          rsadd1.open sqladd1,conn,1,3
          rsadd1("sb_ggxh")=ReplaceBadChar(Trim(Request("gh_xhupdate")))  
      rsadd1("sb_qydate")=ReplaceBadChar(Trim(request("gh_date")))
	  rsadd1.update
          rsadd1.close
          set rsadd1=nothing
	  
	end if 
	
      set rsadd=nothing
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sb where sb_id="&Trim(request("sbid"))
      rsedit.open sqledit,conn,1,3
      	  rsedit("sb_update")=now()
		  if ReplaceBadChar(Trim(request("jx_ylwt")))<>"" then rsedit("sb_whqk")="2"  else rsedit("sb_whqk")="1"
      rsedit.update
      rsedit.close
      set rsedit=nothing
	'rsedit("sbid")=request("sbid")
	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0))(0)
	  Dwt.savesl "设备管理-检修记录-"&sbclassname,"添加",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0)&" 日期："&ReplaceBadChar(Trim(request("jx_date")))
	'Dwt.out"<Script Language=Javascript>history.go(-2)<Script>"
     response.write"<Script Language=Javascript>location.href='?sbid="&sb_id&"&sbclassid="&sbclass_id&"';</Script>"
end sub


sub edit()
    sqledit="SELECT * from sbjx where jx_id="&Trim(Request("jxid"))
	set rsedit=server.createobject("adodb.recordset")
    rsedit.open sqledit,conn,1,1


   Dwt.out"<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'  >"& vbCrLf
   Dwt.out"<form method='post' action='sb_jxjl.asp' name='formadd' onsubmit='javascript:return CheckAdd();'>"& vbCrLf
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.out"<Div align='center'><strong>编辑   "&sb_wh&" 检修记录</strong></Div></td>    </tr>"& vbCrLf
  
  
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修类别： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
    dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxlb where sbjxlb_zclass=0 order by  sbjxlb_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">暂无内容</p>" 
  else
	  do while not rsbody.eof 
						'二级
				sqlz="SELECT * from sbjxlb where sbjxlb_zclass="&rsbody("sbjxlb_id")&" order by  sbjxlb_orderby aSC"& vbCrLf
				set rsz=server.createobject("adodb.recordset")
				rsz.open sqlz,conn,1,1
				if rsz.eof and rsz.bof then 
					dwt.out"<input type='radio' name='jx_lb' value='"&rsbody("sbjxlb_id")&"'" 
							if not isnull(rsedit("jx_lb")) then 
							  if cint(rsedit("jx_lb"))=cint(rsbody("sbjxlb_id")) then 
								 dwt.out " checked "
							  end if    
							end if 
							Dwt.out ">"  	
							dwt.out rsbody("sbjxlb_name") & "<br>"
				else
					do while not rsz.eof
					
						dwt.out"<input type='radio' name='jx_lb' value='"&rsz("sbjxlb_id")&"'" 
							if not isnull(rsedit("jx_lb")) then 
							  if cint(rsedit("jx_lb"))=cint(rsz("sbjxlb_id")) then 
								 dwt.out " checked "
							  end if 
							end if     
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name")&":"&rsz("sbjxlb_name") & "<br>"
					rsz.movenext
					loop
				end if 	
				rsz.close
				set rsz=nothing
			
		rsbody.movenext
		loop
  end if 
  rsbody.close
  set rsbody=nothing
   
   
   
		
  

   dwt.out "</td></tr>"& vbCrLf

  
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>故障现象： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassgz,jxgzname
   sbclassgz=conn.Execute("SELECT sb_jxgzxx_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   if not isnull( sbclassgz ) then 
	  sbclassgz1=split(sbclassgz,",")
	 For i = LBound(sbclassgz1) To UBound(sbclassgz1)
              	jxgzname=getjxgzxx(sbclassgz1(i))

				'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +':'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_gzxx_new' value='"&sbclassgz1(i)&"'"
				call checkbox(rsedit("jx_gzxx_new"),sbclassgz1(i),"")
				dwt.out">"	
				dwt.out jxgzname & "<br>"
				
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_gzxx_new' value='0' onclick='clickgzxxqt()'"
   call checkbox(rsedit("jx_gzxx_new"),0,rsedit("jx_gzxx"))
   dwt.out ">其他<span id=jxgzxxspan></span>"	
			
   dwt.out "</td></tr>"& vbCrLf
   



   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修内容： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassjx,jxnrname
   sbclassjx=conn.Execute("SELECT sb_jxnr_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   
   
   
   dim is999  '用于判断当前值是否包含有更换记录  JS中用
   is999=false
   if not isnull( rsedit("jx_nr_new") ) then 
	  sbclassjx1=split(rsedit("jx_nr_new"),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
              	if cint(sbclassjx1(i))=9999 then is999=true
   	 Next 
	end if  


   if not isnull( sbclassjx ) then 
	  sbclassjx1=split(sbclassjx,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
                jxnrname=getjxnr(sbclassjx1(i))
				'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +':'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassjx1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_nr_new' value='"&sbclassjx1(i)&"' "
				call checkbox(rsedit("jx_nr_new"),sbclassjx1(i),"")

				dwt.out " >"	
				dwt.out jxnrname & "<br>"
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_nr_new' value='0'  id='jxnrqt' onclick='clickjxnrqt()' "
	call checkbox(rsedit("jx_nr_new"),0,rsedit("jx_nr"))
   dwt.out " >其他<span id=jxnrspan></span>"	
   dwt.out"<br><input type='checkbox' name='jx_nr_new' value='9999' onclick='clickjxnrgh()'"
	call checkbox(rsedit("jx_nr_new"),9999,"")
   dwt.out">更换<span id='jxnrgh'></span>"	

   dwt.out "</td></tr>"& vbCrLf
   %>
      <script language="JavaScript">
		
		//检测故障显示的其他是否选中
		function clickgzxxqt(){
		  var checkName = document.getElementsByName ("jx_gzxx_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==0)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			jxgzxxspan.innerHTML="：<input name='jx_gzxx' id='jx_gzxx' type='text'  value='<%=rsedit("jx_gzxx")%>'>"
			}else{
			jxgzxxspan.innerHTML=""
		  }
		}
		//检测检修内容的其他是否选中
		function clickjxnrqt(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==0)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			jxnrspan.innerHTML="：<input name='jx_nr' id='jx_nr' type='text'  value='<%=rsedit("jx_nr")%>'>"
			}else{
			jxnrspan.innerHTML=""
		  }
		}
		//检测检修内容的更换是否选中
		function clickjxnrgh(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//根据组件名获取组建对象
		  var ischecked
		  //循环checkbox，判断是否包含选中项
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//如果有选中项，则返回true
				  
				  if(checkName[i].value==9999)ischecked= true;   //如果选中的是0,也就是"其他"输出真
			  }
		  }   
		  if(ischecked){
			  <%
			  ghqxh=conn.Execute("SELECT sb_ggxh FROM sb  WHERE sb_id="&sb_id)(0)
			  
			  %>
			  <%
			  if is999 then 
			     ghqxh=conn.Execute("SELECT gh_xh FROM sbgh  WHERE jx_id="&Trim(Request("jxid")))(0)
			    'dwt.out "sfsdfdsfd"
			     ghxhupdate=conn.Execute("SELECT gh_xhupdate FROM sbgh  WHERE jx_id="&Trim(Request("jxid")))(0)
			  end if 
			  %>
			  

			jxnrgh.innerHTML="：更换前型号 <input name='gh_xh' id='gh_xh'  type='text' value='<%=ghqxh%>'>&nbsp;更换后型号<input name='gh_xhupdate'  id='gh_xhupdate'  value='<%=ghxhupdate%>' type='text'>"
			}else{
			jxnrgh.innerHTML=""
		  }
		}
      </script>
   <%
   
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修时间： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   Dwt.out"<input name='jx_date'  type='text'  id='jx_date'  class='Wdate' onFocus=""var jx_enddate=$dp.$('jx_enddate');WdatePicker({onpicked:function(){pickedFunc();jx_enddate.focus();},dateFmt:'yyyy/MM/dd HH:mm',maxDate:'#F{$dp.$D(\'jx_enddate\')}'})""   readOnly  value='"&rsedit("jx_date")&"'>"
   dwt.out " 至 "
   
   Dwt.out"<input name='jx_enddate' type='text'  id='jx_enddate'  class='Wdate'   onFocus=""WdatePicker({dateFmt:'yyyy/MM/dd HH:mm',minDate:'#F{$dp.$D(\'jx_date\')}',onpicked:pickedFunc})""   readOnly  value='"&rsedit("jx_enddate")&"'>"
   
   dwt.out "&nbsp;&nbsp;<span id='jxys'></span>  "
   
   Dwt.out"</td></tr>"& vbCrLf
   %>
   <script language="JavaScript">
    // 计算两个日期的间隔天数  
     //   document.all.dateChangDu.value = iDays;
	function pickedFunc(){
		  Date.prototype.dateDiff = function(interval,objDate){    
		//若参数不足或 objDate 不是日期物件则回传 undefined    
		if(arguments.length<2||objDate.constructor!=Date) return undefined;    
		switch (interval) {      
		//计算秒差    
		 // case "s":return parseInt((objDate-this)/1000);      
		  //计算分差    
			case "n":return parseInt(Math.round(((objDate-this)/60000)*100)/100);      
			//计算时差    
			  case "h":return Math.round(((objDate-this)/3600000)*100)/100;      
			  //计算日差      
			 // case "d":return parseInt((objDate-this)/86400000);      
			  //计算月差      
			 // case "m":return (objDate.getMonth()+1)+((objDate.getFullYear()-this.getFullYear())*12)-(this.getMonth()+1);      
			  //计算年差      
			 // case "y":return objDate.getFullYear()-this.getFullYear();      
			
			  //输入有误      
			  default:return undefined;    
			}
		 }
		//document.all.dateChangDu.value = document.all.jx_date.value;
			  var sDT = new Date(document.all.jx_date.value);
			  var eDT = new Date(document.all.jx_enddate.value);
			  jxys.innerHTML=("用时："+ sDT.dateDiff("h",eDT)+"小时");
	
	}

</script>  
   
   <%
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检修负责人： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_fzren' type='text'  value='"&rsedit("jx_fzren")&"' ></td></tr>"& vbCrLf
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检 修 人： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_ren' type='text' value='"&rsedit("jx_ren")&"'><br>两个字的姓名,两字中间请勿添加空格或其他字符 <br>多个检修人,每个人的姓名中间请用空格区分,请勿使用其他字符</td></tr>"& vbCrLf
   
      Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>遗留问题： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   		
					sqlz="SELECT * from sbjxylwt order by  sbjxylwt_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
						do while not rsz.eof
							dwt.out"<input type='checkbox' name='jx_ylwt' value='"&rsz("sbjxylwt_id")&"'" 
							dwt.out rsedit("jx_ylwt")&"-"&rsz("sbjxylwt_id")
					       
						   dwt.out checkbox(rsedit("jx_ylwt"),rsz("sbjxylwt_id"),"")
								
							Dwt.out ">"	
							dwt.out rsz("sbjxylwt_name") & "<br>"
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing 
  

   dwt.out "<b>上述都不选择表示无遗留问题</b>"& vbCrLf
   dwt.out "<br><b>如果选择上述项目则表示有遗留问题,设备""完好""状态会更新为不完好</b></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备    注： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_bz' type='text' value='"&rsedit("jx_bz")&"'></td></tr>"& vbCrLf
   
   
   Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
   Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>   <input name='sbid' type='hidden'  value='"&Trim(Request("sbid"))&"'> <input name='sbclassid' type='hidden'  value='"&Trim(Request("sbclassid"))&"'>  <input type='hidden' name='jxid' value='"&Trim(Request("jxid"))&"'>     <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick=""window.location.href='sb_jxjl.asp?sbid="&Trim(Request("sbid"))&"&sbclassid="&Trim(Request("sbclassid"))&"';"" style='cursor:hand;'></td>  </tr>"

   Dwt.out"</form></table>"& vbCrLf























   rsedit.close
   set rsedit=nothing
end sub

'函数名称：checkbox 页面是否选择
'作用：判断检修记录的内容和设备分类的检修内容是否对应 检修记录中和设备分类中的检修内容对应则则输出checked

'jx_gzxx_new  数组,检修记录中保存的值
'sbclassjxid  调用的时候已经分割开的  设备分类中的检修记录内容
'jx_gzxx  兼容旧的数据,如果有此信息,则"其他"是默认选中的
Function checkbox(jx_gzxx_new,sbclassjxid,jx_gzxx)
	dim sbclassjx1,i
	if not isnull( jx_gzxx_new ) then 
	  sbclassjx1=split(jx_gzxx_new,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=cint(sbclassjxid) then dwt.out " checked "
	 Next 
	 
	end if  
	 if jx_gzxx<>"" then dwt.out " checked "
end Function


sub saveedit()
	jxid=ReplaceBadChar(Trim(request("jxID")))
	
	'编辑保存
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sbjx where jx_ID="&jxid
	
	rsedit.open sqledit,conn,1,3







	rsedit("jx_lb")=ReplaceBadChar(Trim(Request("jx_lb")))
	rsedit("jx_gzxx")=ReplaceBadChar(Trim(Request("jx_gzxx")))
	rsedit("jx_nr")=ReplaceBadChar(Trim(request("jx_nr")))
	rsedit("jx_gzxx_new")=ReplaceBadChar(Trim(Request("jx_gzxx_new")))
	rsedit("jx_nr_new")=ReplaceBadChar(Trim(request("jx_nr_new")))
	rsedit("jx_date")=Trim(request("jx_date"))
	rsedit("jx_enddate")=Trim(request("jx_enddate"))
	rsedit("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsedit("jx_ren")=ReplaceBadChar(Trim(request("jx_ren")))
	rsedit("jx_ylwt")=ReplaceBadChar(Trim(request("jx_ylwt")))
	rsedit("jx_bz")=ReplaceBadChar(Trim(request("jx_bz")))
	rsedit("sb_id")=ReplaceBadChar(Trim(request("sbid")))
	
      rsedit.update
      rsedit.close
      set rsedit=nothing
	
	ghxh=ReplaceBadChar(Trim(request("gh_xh")))
	ghxhupdate=ReplaceBadChar(Trim(request("gh_update")))


   dim isgh '是否选中更换
  isgh =false
   if not isnull( ReplaceBadChar(Trim(request("jx_nr_new"))) ) then 
	  sbclassjx1=split(ReplaceBadChar(Trim(request("jx_nr_new"))),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=9999 then isgh=true
	 Next 
	 
	end if  

	'如果选中了更换数据,则保存更信息,如果没选中 则判断更换记录里有没有些JXID的数据,有的话删除
	if isgh and jxid<>"" then 
 
                    '检测是否有些JXID的更换记录,没有就增加,有就更新
					sqlz="SELECT * from sbgh where jx_id="&jxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
						'保存新增更换记录
						set rsaddgh=server.createobject("adodb.recordset")
						sqladdgh="select * from sbgh" 
						rsaddgh.open sqladdgh,conn,1,3
						rsaddgh.addnew
						rsaddgh("jx_id")=jxid
						rsaddgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
						rsaddgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
					   rsaddgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
					   rsaddgh.update
						rsaddgh.close
						set rsaddgh=nothing
					else
					  set rseditgh=server.createobject("adodb.recordset")
					  sqleditgh="select * from sbgh where jx_ID="&jxid
					  
					  rseditgh.open sqleditgh,conn,1,3
				  
						rseditgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
						rseditgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
					   'rseditgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
				  
						rseditgh.update
						rseditgh.close
						set rsedight=nothing

					end if 	
					rsz.close
					set rsz=nothing 


	 
	 
	  
			 '更改设备档案中相应位号设备的规格型号2008-9-19
			set rsadd1=server.createobject("adodb.recordset")
				sqladd1="select * from sb where sb_id="&Trim(request("sbid"))
				rsadd1.open sqladd1,conn,1,3
				rsadd1("sb_ggxh")=ReplaceBadChar(Trim(Request("gh_xhupdate")))  
			rsadd1("sb_qydate")=ReplaceBadChar(Trim(request("gh_date")))
			rsadd1.update
				rsadd1.close
				set rsadd1=nothing
	  
	else
'删除更换 信息
					sqlz="SELECT * from sbgh where jx_id="&jxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
					  set rsdel=server.createobject("adodb.recordset")
					  sqldel="delete * from sbgh where jx_id="&jxid
					  rsdel.open sqldel,conn,1,3
					  'rsdel.close
					  set rsdel=nothing  
					 
					 

					end if 	
					rsz.close
					set rsz=nothing 
	end if 
	
	   
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sb where sb_id="&Trim(request("sbid"))
      rsedit.open sqledit,conn,1,3
      	  rsedit("sb_update")=now()
		  if ReplaceBadChar(Trim(request("jx_ylwt")))<>"" then rsedit("sb_whqk")="2"  else rsedit("sb_whqk")="1"
      rsedit.update
      rsedit.close
      set rsedit=nothing



















	

	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0))(0)
	  Dwt.savesl "设备管理-检修记录-"&sbclassname,"编辑",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0)&" 日期："&ReplaceBadChar(Trim(request("jx_date")))

	Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
	jx_ID=request("jxID")
	sb_id=conn.Execute("SELECT sb_id FROM sbjx WHERE jx_id="&jx_id)(0)
	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&sb_id)(0))(0)
	deljxdate=conn.Execute("SELECT jx_date FROM sbjx WHERE jx_id="&jx_id)(0)
	  Dwt.savesl "设备管理-检修记录-"&sbclassname,"删除",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&sb_id)(0)&" 时间："&deljxdate
	
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from sbjx where jx_id="&jx_id
	rsdel.open sqldel,conn,1,3
	Dwt.out"<Script Language=Javascript>history.back()</Script>"
	'rsdel.close
	set rsdel=nothing  

end sub


'**********************************************
'登录名以区分是否有修改权限页面jxjl.asp
'******************************8
sub jxxuanxiang(id,sb_id,sb_sscj)
 if session("levelclass")=sb_sscj or session("levelclass")=0 then 
	Dwt.out"<a href=sb_jxjl.asp?action=edit&sbid="&sb_id&"&sbclassid="&sbclass_id&"&jxid="&rsjx("jx_id")&">编</a>&nbsp;"
	Dwt.out"<a href=sb_jxjl.asp?action=del&jxid="&rsjx("jx_id")&"&sbclassid="&sbclass_id&"&sbid="&sb_id&" onClick=""return confirm('确定要删除此记录吗？');"">删</a>"
 else
    Dwt.out"&nbsp;"
 end if 
end sub



Call CloseConn
%>