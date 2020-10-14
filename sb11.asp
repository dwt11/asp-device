<%@language=vbscript codepage=936 %>
<%
'此页为分页的另一种形式，但只能按SB_ID DESC排序
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
dim starttime : starttime=timer
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<%
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh")) 
sb_classid = Trim(Request("sbclassid"))
if sb_classid="" then sb_classid=1
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sb_classid&" and sbclass_zclass=0")(0)

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title> 技术档案管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out"<script type=""text/javascript"" src=""js/ajax.js""></script>"&vbcrlf
dwt.out"<script type=""text/javascript"" src=""js/common.js""></script>"&vbcrlf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out " if(document.form.sb_sscj.value==''){" & vbCrLf
dwt.out "      alert('请选择所属车间！');" & vbCrLf
dwt.out "   document.form.sb_sscj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out " if(document.form.sb_ssgh.value==0){" & vbCrLf
dwt.out "      alert('请选择所属装置！');" & vbCrLf
dwt.out "   document.form.sb_ssgh.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out " if(document.form.sb_wh.value==''){" & vbCrLf
dwt.out "      alert('请添写位号！');" & vbCrLf
dwt.out "   document.form.sb_wh.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out " if(document.form.sb_sccj.value==''){" & vbCrLf
dwt.out "      alert('请添写生产厂家！');" & vbCrLf
dwt.out "   document.form.sb_sccj.focus();" & vbCrLf
dwt.out "      return false;" & vbCrLf
dwt.out "    }" & vbCrLf

dwt.out "    }" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
action=request("action")
select case action
  case "add"
      call add'添加设备分类选择
  case "addsb"
      call addsb'选择分类后添加设备页面
  case "saveaddsb"
      call saveaddsb'设备添加保存
  case "edit"
      call edit
  case "saveedit"'编辑子分类
      call saveedit'编辑保存子分类
  case "del"
      call del     '删除子分类信息
  case ""
      call main
end select	  	 

sub add()
   dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   dwt.out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   dwt.out"<div align='center'><strong>添加新设备</strong></div></td></tr>"& vbCrLf
    dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>设备分类： </strong></td>"
	dwt.out"<td width='88%' class='tdbg'><select name='sb_dclass' size='1' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">"
	call formdclass
	dwt.out"</select>"
	dwt.out"</td></tr></table>"
end sub


sub addsb()
'sbclass_id=request("sbclassid")
	dwt.out"<form method='post' action='sb.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out"<tr class='title'>"& vbCrLf
	dwt.out"<td height='22' colspan='2'><div align=center><strong>新增 "&sb_classname&" 设备</strong></div></tr>"& vbCrLf
	
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	dwt.out"<select name='sb_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select>"  	 & vbCrLf
    dwt.out "<select name='sb_ssgh' size='1' >" & vbCrLf
    dwt.out "<option  selected>选择装置分类</option>" & vbCrLf
    dwt.out "</select></td></tr>  "  & vbCrLf
    dwt.out "<script><!--" & vbCrLf
    dwt.out "var groups=document.form.sb_sscj.options.length" & vbCrLf
    dwt.out "var group=new Array(groups)" & vbCrLf
    dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    dwt.out "group[i]=new Array()" & vbCrLf
    dwt.out "group[0][0]=new Option(""选择装置分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from ghname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   dwt.out "group["&rscj("levelid")&"][0]=new Option(""无装置分类"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("gh_name")&""","""&rsbz("ghid")&""");" & vbCrLf
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




    dwt.out "var temp=document.form.sb_ssgh" & vbCrLf
    dwt.out "function redirect(x){" & vbCrLf
    dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    dwt.out "temp.options[m]=null" & vbCrLf
    dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    dwt.out "}" & vbCrLf
    dwt.out "temp.options[0].selected=true" & vbCrLf
    dwt.out "}//--></script>" & vbCrLf



  else 	 
   dwt.out"<input name='sb_sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
   dwt.out"<input name='sb_sscj' type='hidden' value="&session("level")&">"& vbCrLf
   sql="SELECT * from ghname where sscj="&session("level")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   dwt.out"<select name='sb_ssgh' size='1'>"
   
   if rs.eof and rs.bof then 
   	  dwt.out"<option value='0'>未添加装置</option>"
   else   
	  'dwt.out"<option value='0'>车间</option>"
      do while not rs.eof
	     dwt.out"<option value='"&rs("ghid")&"'>"&rs("gh_name")&"</option>"
	  rs.movenext
      loop
	  end if 
	 dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
    dwt.out"</td></tr>  "  	 

	
	
	if zclassor(sb_classid) then
		dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>类型： </strong></td>"   & vbCrLf   
		dwt.out"<td width='70%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass(sb_classid)
		dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>位号： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_wh' type='text' ></td></tr>"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>设备特性： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out" <label><input type='checkbox' name='sb_isls'/>是否联锁 </label>"
	dwt.out" <label><input type='checkbox' name='sb_iszj'/>是否周检 </label>"
	dwt.out" <label><input type='checkbox' name='sb_isbw'/>是否保温 </label>"
	dwt.out" <label><input type='checkbox' name='sb_isjl'/>是否计量核算 </label>"
	
	dwt.out "</td></tr>"& vbCrLf
		dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>完好： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_whqk' size='1' >"
	dwt.out"<option value='0'"
	
	'if rsedit("sb_whqk")=0 then dwt.out" selected" 
	dwt.out">请选择完好情况</option>"
	dwt.out"<option value='1' "
	'if rsedit("sb_whqk")=1 then dwt.out"selected"
	dwt.out">完好</option>"
	dwt.out"<option value='2'"
	'if rsedit("sb_whqk")=2 then dwt.out"selected"
	dwt.out" >不完好</option>"
	dwt.out"</select></td></tr>"
	
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>准确： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_zqqk' size='1' >"
	dwt.out"<option value='0'>请选择准确情况</option>"
	dwt.out"<option value='1' >最大最小</option>"
	dwt.out"<option value='2'>中间</option>"
	dwt.out"<option value='3'>>95%</option>"
	dwt.out"</select></td></tr>"

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>投运： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_tyqk' size='1' >"
	dwt.out"<option value='0'>请选择投运情况</option>"
	dwt.out"<option value='1'>投运</option>"
	dwt.out"<option value='2'>原因未投运</option>"
	dwt.out"<option value='3'>工艺原因未投运</option>"
	dwt.out"</select></td></tr>"

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>分级： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_fj' size='1' >"
	dwt.out"<option value='0'>请选择分级</option>"
	dwt.out"<option value='1'>一级</option>"
	dwt.out"<option value='2'>二级</option>"
	dwt.out"<option value='3'>三级</option>"
	dwt.out"</select></td></tr>"
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>规格型号： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_ggxh' type='text'></td></tr>"& vbCrLf
	
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			'字段名
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'字段在页面中显示的名称
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	

	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
	sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  '去除最右边逗号
	sb_tablename=split(sb_tablename,",")
	sb_tablebody=split(sb_tablebody,",")
	
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
		dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"： </strong></td>"   & vbCrLf   
	    dwt.out"<td width='70%' class='tdbg'><input name='"&sbtablename&"' type='text'></td></tr>"& vbCrLf
	next
	

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>生产厂家： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_sccj' type='text'></td></tr>"& vbCrLf
   dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>启用时间： </strong></td>"   & vbCrLf   
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='sb_qydate' type='text' value="&date()&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this,sb_qydate, ' yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>备注： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_whbeizhu' type='text'></td></tr>"& vbCrLf
	
	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveaddsb'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>     <input  type='submit' name='Submit' value=' 添   加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub

sub saveaddsb()

'新增保存
	sb_classid=request("sbclassid")
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sb"
	rsadd.open sqladd,conn,1,3
	      rsadd.addnew
on error resume next
    rsadd("sb_dclass")=ReplaceBadChar(Trim(Request("sbclassid")))
	rsadd("sb_sscj")=ReplaceBadChar(Trim(Request("sb_sscj")))
	rsadd("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsadd("sb_dclass")) then 	rsadd("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '判断是否有子分类,再保存
	rsadd("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsadd("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsadd("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsadd("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsadd("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
	rsadd("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	
	
	    sb_isls=request("sb_isls")
	if sb_isls="on" then
	  sb_isls=true
	else
	  sb_isls=false
	end if  
	rsadd("sb_isls")=sb_isls
    
	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsadd("sb_iszj")=sb_iszj
    
	sb_isbw=request("sb_isbw")
	if sb_isbw="on" then
	  sb_isbw=true
	else
	  sb_isbw=false
	end if  
	rsadd("sb_isbw")=sb_isbw
    
	sb_isjl=request("sb_isjl")
	if sb_isjl="on" then
	  sb_isjl=true
	else
	  sb_isjl=false
	end if  
	rsadd("sb_isjl")=sb_isjl

	
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	
	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
	sb_tablename=split(sb_tablename,",")
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
        rsadd(sbtablename)=ReplaceBadChar(Trim(request(sbtablename)))
	next
	
	rsadd("sb_sccj")=ReplaceBadChar(Trim(request("sb_sccj")))
	rsadd("sb_bz")=ReplaceBadChar(Trim(request("sb_bz")))
	rsadd.update
	rsadd.close
	dwt.out"<Script Language=Javascript>location.href='sb.asp?sbclassid="&sb_classid&"'</Script>"

end sub


sub saveedit()
'编辑保存
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sb where sb_ID="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,conn,1,3
	on error resume next

	rsedit("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsedit("sb_dclass")) then 	rsedit("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '判断是否有子分类,再保存
	rsedit("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsedit("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsedit("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsedit("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsedit("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
    rsedit("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	sb_isls=request("sb_isls")
	if sb_isls="on" then
	  sb_isls=true
	else
	  sb_isls=false
	end if  
	rsedit("sb_isls")=sb_isls
    
	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsedit("sb_iszj")=sb_iszj
    
	sb_isbw=request("sb_isbw")
	if sb_isbw="on" then
	  sb_isbw=true
	else
	  sb_isbw=false
	end if  
	rsedit("sb_isbw")=sb_isbw
    
	sb_isjl=request("sb_isjl")
	if sb_isjl="on" then
	  sb_isjl=true
	else
	  sb_isjl=false
	end if  
	rsedit("sb_isjl")=sb_isjl

	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	
	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
	sb_tablename=split(sb_tablename,",")
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
        rsedit(sbtablename)=ReplaceBadChar(Trim(request(sbtablename)))
	next
	
	rsedit("sb_sccj")=ReplaceBadChar(Trim(request("sb_sccj")))
	rsedit("sb_bz")=ReplaceBadChar(Trim(request("sb_bz")))
	rsedit.update
	rsedit.close
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub edit()
	sb_id=ReplaceBadChar(Trim(request("id")))
	'sb_classid=Trim(request("sbclassid"))
	'sb_wh=conn.Execute("SELECT sb_wh FROM sb WHERE  sb_id="&sb_id)(0)

	sqledit="SELECT * from sb where sb_id="&sb_id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,conn,1,1
	dwt.out"<form method='post' action='sb.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	dwt.out"<tr class='title'>"& vbCrLf
	dwt.out"<td height='22' colspan='2'><div align=center><strong>"&sb_classname&"编辑 "
	dwt.out"位号:"& vbCrLf
	dwt.out rsedit("sb_wh")&"</strong></div></tr>"& vbCrLf
	
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sb_sscj"))&"'></td></tr>"& vbCrLf
    dwt.out"<input name='sb_sscj' type='hidden' value="&rsedit("sb_sscj")&">"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>所属装置： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_ssgh' size='1' >"
	call formgh (rsedit("sb_ssgh"),rsedit("sb_sscj"))
	dwt.out"</select></td></tr>"
	
	
	if zclassor(rsedit("sb_dclass")) then
		dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>类型： </strong></td>"   & vbCrLf   
		dwt.out"<td width='70%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass(rsedit("sb_zclass"))
		dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>位号： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_wh' type='text' value='"&rsedit("sb_wh")&"'></td></tr>"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>设备特性： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out" <label><input type='checkbox' name='sb_isls' "
	if rsedit("sb_isls") then dwt.out "checked='checked'"
	dwt.out" />是否联锁 </label>"
	dwt.out" <label><input type='checkbox' name='sb_iszj' "
	if rsedit("sb_iszj") then dwt.out "checked='checked'"
	dwt.out" />是否周检 </label>"
	dwt.out" <label><input type='checkbox' name='sb_isbw' "
	if rsedit("sb_isbw") then dwt.out "checked='checked'"
	dwt.out" />是否保温 </label>"
	dwt.out" <label><input type='checkbox' name='sb_isjl' "
	if rsedit("sb_isjl") then dwt.out "checked='checked'"
	dwt.out" />是否计量核算 </label>"
	
	dwt.out "</td></tr>"& vbCrLf


		dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>完好： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_whqk' size='1' >"
	dwt.out"<option value='0'"
	
	if rsedit("sb_whqk")=0 then dwt.out" selected" 
	dwt.out">请选择完好情况</option>"
	dwt.out"<option value='1' "
	if rsedit("sb_whqk")=1 then dwt.out"selected"
	dwt.out">完好</option>"
	dwt.out"<option value='2'"
	if rsedit("sb_whqk")=2 then dwt.out"selected"
	dwt.out" >不完好</option>"
	dwt.out"</select></td></tr>"
	
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>准确： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_zqqk' size='1' >"
	dwt.out"<option value='0'"
	if rsedit("sb_zqqk")=0 then dwt.out" selected" 
	dwt.out">请选择准确情况</option>"
	dwt.out"<option value='1' "
	if rsedit("sb_zqqk")=1 then dwt.out"selected"
	dwt.out">最大最小</option>"
	dwt.out"<option value='2'"
	if rsedit("sb_zqqk")=2 then dwt.out"selected"
	dwt.out" >中间</option>"
	dwt.out"<option value='3'"
	if rsedit("sb_zqqk")=3 then dwt.out"selected"
	dwt.out" >>95%</option>"
	dwt.out"</select></td></tr>"

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>投运： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_tyqk' size='1' >"
	dwt.out"<option value='0'"
	if rsedit("sb_tyqk")=0 then dwt.out" selected" 
	dwt.out">请选择投运情况</option>"
	dwt.out"<option value='1' "
	if rsedit("sb_tyqk")=1 then dwt.out"selected"
	dwt.out">投运</option>"
	dwt.out"<option value='2'"
	if rsedit("sb_tyqk")=2 then dwt.out"selected"
	dwt.out" >原因未投运</option>"
	dwt.out"<option value='3' "
	if rsedit("sb_tyqk")=3 then dwt.out"selected"
	dwt.out">工艺原因未投运</option>"
	dwt.out"</select></td></tr>"

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>分级： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'>"
	dwt.out"<select name='sb_fj' size='1' >"
	dwt.out"<option value='0'"
	if rsedit("sb_fj")=0 then dwt.out" selected" 
	dwt.out">请选择分级</option>"
	dwt.out"<option value='1' "
	if rsedit("sb_fj")=1 then dwt.out"selected"
	dwt.out">一级</option>"
	dwt.out"<option value='2'"
	if rsedit("sb_fj")=2 then dwt.out"selected"
	dwt.out" >二级</option>"
	dwt.out"<option value='3' "
	if rsedit("sb_fj")=3 then dwt.out"selected"
	dwt.out">三级</option>"
	dwt.out"</select></td></tr>"
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>规格型号： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_ggxh' type='text' value='"&rsedit("sb_ggxh")&"'></td></tr>"& vbCrLf
	
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			'字段名
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'字段在页面中显示的名称
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
			
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	

	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
	sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  '去除最右边逗号
	sb_tablename=split(sb_tablename,",")
	sb_tablebody=split(sb_tablebody,",")
	
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
		
'		if mid(sbtablename,4,1)="b" then
'		
'		'BOOL类型字段,以第一个字为真,第二个字为假,例"气电类型" 气真,电假
'			dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"： </strong></td>"   & vbCrLf   
'			dwt.out"<td width='70%' class='tdbg'>"
'			dwt.out"<select name='sbtablename' size='1' >"
'			dwt.out"<option value='0'"
'			if rsedit(sbtablename)=0 then dwt.out" selected" 
'			dwt.out">请选择"&sb_tablebody(sb_tablei)&"</option>"
'			dwt.out"<option value='true' "
'			if rsedit(sbtablename)=true then dwt.out"selected"
'			dwt.out">"&left(sb_tablebody(sb_tablei),1)&"</option>"
'			dwt.out"<option value='false'"
'			if rsedit(sbtablename)=false then dwt.out"selected"
'			dwt.out" >"&mid(sb_tablebody(sb_tablei),2,1)&"</option>"
'			dwt.out"</select></td></tr>"
'		else 
			dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"： </strong></td>"   & vbCrLf   
			dwt.out"<td width='70%' class='tdbg'><input name='"&sbtablename&"' type='text' value='"&rsedit(sbtablename)&"'></td></tr>"& vbCrLf
	   'end if 
		'dwt.out sbtablename&"<br>"&sb_tablei
   'dwt.out sb_tablename(sb_tablei)&" "&sb_tablebody(sb_tablei)
	next
	

	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>生产厂家： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_sccj' type='text' value='"&rsedit("sb_sccj")&"'></td></tr>"& vbCrLf
	
   dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>启用时间： </strong></td>"   & vbCrLf   
   dwt.out"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out"<input name='sb_qydate' type='text' value="&rsedit("sb_qydate")&">"
   dwt.out"<a href='#' onClick=""popUpCalendar(this,sb_qydate, ' yyyy-mm-dd'); return false;"">"
   dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	dwt.out"<tr class='tdbg'><td width='30%' align='right' class='tdbg'><strong>备注： </strong></td>"   & vbCrLf   
	dwt.out"<td width='70%' class='tdbg'><input name='sb_whbeizhu' type='text' value='"&rsedit("sb_bz")&"'></td></tr>"& vbCrLf
	
	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>   <input name='id' type='hidden'  value='"&Trim(Request("id"))&"'> <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	
	dwt.out"</table></form>"
  rsedit.close
  set rsedit=nothing
  conn.close
  set conn=nothing

end sub
sub main()
'    url="sb.asp?sbclassid="&sb_classid
'	if request("search")="sscjs" then  url="sb.asp?sscj="&sscjid&"&search=sscjs&sbclassid="&sb_classid
'	if request("search")="ssghs" then  url="sb.asp?ssgh="&ssghid&"&search=ssghs&sbclassid="&sb_classid
'	if request("search")="keys" then  url="sb.asp?keyword="&keys&"&search=keys&sbclassid="&sb_classid

	if request("search")="sscjs" then title=" 搜索车间 "&sscjh(sscjid) 
	if request("search")="ssghs" then title=" 搜索装置 "&gh(ssghid) 
	if request("search")="keys" then title=" 搜索位号 '"&keys&" '"

	dwt.out "<table width=100% border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
	dwt.out " <tr class='topbg'>"& vbCrLf
	dwt.out "   <td height='22' colspan='2' align='center'><strong>设备档案－－"&sb_classname&title&"</strong></td>"& vbCrLf
	dwt.out "  </tr>  "& vbCrLf
	call search	()
	dim v1 ,v2,v3,totall,zh,v4
	v1= "<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj=1")(0)&"</font>" 
	
	v2= "<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj=2")(0)&"</font>" 
	
	v3= "<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj=3")(0)&"</font>" 
	v4="<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj=4")(0)&"</font>" 
	
	zh="<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj=5")(0)&"</font>" 
	
	
	totall= "<font color='#006600'>"&conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid)(0)&"</font>" 
	dwt.out "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"& vbCrLf
	dwt.out " <tr >"& vbCrLf
	dwt.out "   <td height='22' colspan='2' align='center'>  <strong>维一："&v1&"</strong>   <strong>维二："&v2&"</strong>     <strong>维三："&v3&"</strong>     <strong>维四："&v4&"</strong>     <strong>综合："&zh&"</strong>     <strong>合计："&totall&"</strong></td>"& vbCrLf
	dwt.out "  </tr>  "& vbCrLf
	dwt.out "</table>"& vbCrLf
	sqlbody=""
	sqlbody=" where sb_dclass="&sb_classid 
	if request("search")="sscjs" then sqlbody=sqlbody&" and sb_sscj="&sscjid
	if request("search")="ssghs" then sqlbody=sqlbody&" and sb_ssgh="&ssghid
	if request("search")="keys" then sqlbody=sqlbody&" and sb_wh  like '%" &keys& "%' "
	set sb=new dwt_page '实例化一个类
	pgsz=request("pgsz")
	if pgsz="" then pgsz=20
	sb.getconn=conn     '必须，指定adodb.connection数据库连接对象
	sb.dwttable="sb"   '必须，要进行分页的数据库表
	sb.dwtsql=sqlbody   '必须，SQL查询条件
	sb.dwtpagesize=pgsz   '可选，指定每页显示的记录数，默认值为10
	sb.dwtpagestyle=1'可选，显示分页导舰的样式，取值为1或2，默认值为1
	'sb.dwtorderby="sb_sscj asc,sb_ssgh asc,sb_wh asc"
	sb.dwtorderby="sb_id desc"
	sb.dwt_set   '必须，根据上面的指定的设置，执行分页程序
	
	dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
	dwt.out "<tr class=""title"">"
	dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>序号</strong></div></td>"
	dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>车间</strong></div></td>"
	dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>装置</strong></div></td>"
	if zclassor(sb.dwtrs("sb_dclass")) then 	dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>类型</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>位 号</strong></div></td>"
	dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>完好</strong></div></td>"
	dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>准确</strong></div></td>"
	dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>投运</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>分级</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>规格型号</strong></div></td>"
	'取字段的名称
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			dwt.out "<td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>"&rsbody1("sbtable_body")&"</strong></div></td>"
			rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	
	
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>生产厂家</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>启用时间</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备 注</strong></div></td>"
	dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选 项</strong></div></td>"
	dwt.out "    </tr>"
		
  for i=1 to sb.dwtpagesize
	if sb.dwtrs.eof then exit for '当到了尾记录的时候，退出for循环
			dwt.out " <tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb.dwtrs("sb_id")&"</div></td>"
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sscjh_d(sb.dwtrs("sb_sscj"))&"</div></td>"
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&GH(sb.dwtrs("sb_ssGH"))&"</div></td>"
			if zclassor(sb.dwtrs("sb_dclass")) then 	dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&zclass(sb.dwtrs("sb_zclass"))&"</div></td>"
			
			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"""
		   if now()-sb.dwtrs("sb_update")<7 then   dwt.out "bgcolor=""#FFFF00"""
		   dwt.out ">"&searchH(uCase(sb.dwtrs("sb_wh")),keys)&"&nbsp;</td>"

			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb_whd(sb.dwtrs("sb_whqk"))&"</div></td>"
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb_zqd(sb.dwtrs("sb_ZQqk"))&"</div></td>"
			dwt.out " <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb_tyd(sb.dwtrs("sb_tyqk"))&"</div></td>"
			dwt.out "  <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&fj(sb.dwtrs("sb_fj"))&"&nbsp;</div></td>"
			dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"">"&sb.dwtrs("sb_ggxh")&"&nbsp;</td>"
		
			'取字段内容
			sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
			set rsbody1=server.createobject("adodb.recordset")
			rsbody1.open sqlbody1,conn,1,1
			if rsbody1.eof and rsbody1.bof then 
				dwt.out "<p align=""center"">暂无内容</p>" 
			else
				do while not rsbody1.eof
				  sbtable_name=rsbody1("sbtable_name")   '取得表的名称
				  dwt.out "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb.dwtrs(sbtable_name)&"</div></td>"
				rsbody1.movenext
				loop
			end if 
			set rsbody1=nothing	
			
			dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"">"&sb.dwtrs("sb_sccj")&"&nbsp;</td>"
			dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"">"&sb.dwtrs("sb_qydate")&"&nbsp;</td>"
			dwt.out " <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sb.dwtrs("sb_bz")&"&nbsp;</div></td>"
			dwt.out "<td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
			dwt.out"  <a href=sb_jxjl.asp?sbid="&sb.dwtrs("sb_id")&"&sbclassid="&sb_classid&">检修</a>&nbsp;<a href=sb_ghjl.asp?sbid="&sb.dwtrs("sb_id")&"&sbclassid="&sb_classid&">更换</a>&nbsp;"
			call sbeditdel(sb.dwtrs("sb_id"),sb.dwtrs("sb_sscj"),"sb.asp?action=edit&sbclassid="&sb_classid&"&id=","sb.asp?action=del&id=",sb_classid)'检修、更换、编辑、删除
			'if session("level")=5 and sb_class=27 then 
			dwt.out "</div></td></tr>"
		sb.dwtrs.movenext '只读向前记录集指针下移

	next
   dwt.out "</table>"&sb.dwtshowpage
end sub
	'dwt.out "程序执行用时：" & timer-starttime
dwt.out "</body></html>"

sub search()
	dim rscj,sqlcj
	dwt.out "<table width=100% border='0' cellpadding='0' cellspacing='0' class='border'>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='sb.asp'>" & vbCrLf
	dwt.out "  <tr class='tdbg'>   <td>" & vbCrLf
	dwt.out "  <font color='0066CC'>位号搜索：</font>" & vbCrLf
	dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50' onFocus='this.select();' autocomplete=""off"" value="&keys&">" & vbCrLf
	dwt.out "  <input type='submit' name='Submit'  value='搜索'>" & vbCrLf
	dwt.out "  <input type='hidden' name='search' value='keys'>" & vbCrLf
	dwt.out "  <input type='hidden' name='sbclassid' value='"&sb_classid&"'>" & vbCrLf
	dwt.out "</td></form><td ><font color='0066CC'>查看所属车间的相关内容：</font>"
	
	dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='sb.asp?search=sscjs&sbclassid="&sb_classid&"&sscj="&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf



	dwt.out "<font color='0066CC'> 所属装置内容：</font>"
	dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>按装置跳转至…</option>" & vbCrLf
	sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
       	dwt.out"<option value='sb.asp?search=ssghs&sbclassid="&sb_classid&"&ssgh="&rsgh("ghid")&"'>"&rsgh("gh_name")&"("&Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_ssgh="&rsgh("ghid")&"and sb_dclass="&sb_classid)(0)&")</option>"& vbCrLf
	
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	dwt.out "     </select>	" & vbCrLf
	dwt.out "</tr></table>" & vbCrLf
end sub

'选项（编辑、删除）
sub sbeditdel(id,sscj,editurl,delurl,sb_classid)
 if session("level")=sscj or session("level")=0 then 
    dwt.out "<a href="&editurl&id&">编辑</a>&nbsp;"
	dwt.out "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 else
    dwt.out "&nbsp;"
 end if 
 if session("level")=5 and sb_classid=27 then 
    dwt.out "<a href="&editurl&id&">编辑</a>&nbsp;"
	dwt.out "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 end if 
end sub

'取子分类名称
function zclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclass= "未编辑" 
  else
     zclass=rsbody("sbclass_name")
  end if
end function

'判断是否有子分类
function zclassor(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_zclass="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclassor=false 
  else
     zclassor=true
  end if
end function


'父分类列表显示
function formdclass()
	dim sqldclass,rsdclass
	set rsdclass=server.createobject("adodb.recordset")
	rsdclass.open sqldclass,conn,1,1
	if rsdclass.eof then 
		dclass="没有任何分类"
	else
		dwt.out"<option value='0'"
		if dclassid=0 then dwt.out " selected" 
			dwt.out">请选择要添加设备的分类</option>"
		do while not rsdclass.eof
			dwt.out"<option value='sb.asp?action=addsb&sbclassid="&rsdclass("sbclass_id")&"'>"&rsdclass("sbclass_name")&"</option>"  & vbCrLf   
		rsdclass.movenext
		loop
	end if 
	rsdclass.close
	set rsdclass=nothing
end function


'子分类列表显示
function formzclass(zclassid)
	dim sqlzclass,rszclass
	if isnull(zclassid) then zclassid=0
		sqlzclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_zclass="&zclassid
	set rszclass=server.createobject("adodb.recordset")
	rszclass.open sqlzclass,conn,1,1
	if rszclass.eof then 
		zclass="未编辑"
	else
		dwt.out"<option value='0'"
		if zclassid=0 then dwt.out " selected" 
			dwt.out">请选择类型</option>"
		do while not rszclass.eof
			dwt.out"<option value='"&rszclass("sbclass_id")&"' "
			if zclassid=rszclass("sbclass_id") then dwt.out "selected"
			dwt.out">"&rszclass("sbclass_name")&"</option>"  & vbCrLf   
		rszclass.movenext
		loop
	end if 
	rszclass.close
	set rszclass=nothing
end function

'完好情况显示
Function sb_whd(whnumb)
	if isnull(whnumb) or whnumb=0 then 
	  sb_whd="未编辑"
	else
		if whnumb=1 then sb_whd="<font color='##006600'>*</font>"  '完好绿
		if whnumb=2 then sb_whd="<font color='#ff0000'>*</font> "	  '不完好红
	end if 
end Function 

'准确情况显示
Function sb_zqd(zqnumb)

	if isnull(zqnumb) or zqnumb=0 then 
	  sb_zqd="未编辑"
	else
		if zqnumb=3 then sb_zqd="***"'>95%
		if zqnumb=2 then sb_zqd="**"		  '中间  
		if zqnumb=1 then sb_zqd="*"  '最大最小
	end if 
end Function 

'投运情况显示
Function sb_tyd(tynumb)

	if isnull(tynumb) or tynumb=0 then 
	  sb_tyd="未编辑"
	else
		if tynumb=1 then sb_tyd="<font color='##006600'>*</font>"   '绿投运
		if tynumb=2 then sb_tyd="<font color='#0000ff'>*</font>"   '蓝仪原因未投运
		if tynumb=3 then sb_tyd="<font color='#ff0000'>*</font>"    '红工艺原因未投运
	end if 
end Function 
Call CloseConn

%>