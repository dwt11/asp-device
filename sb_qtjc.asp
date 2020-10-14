<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
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
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh")) 
sb_classid = Trim(Request("sbclassid"))
   if sb_classid="" then sb_classid=164
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sb_classid)(0)

Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>技术档案管理页</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/grid.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/docs.css' rel='stylesheet' type='text/css'>"& vbCrLf

Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")
select case action
  case "img"
    dwt.out "<br/><Div align=center><b>"&conn.Execute("SELECT sb_wh FROM sbqt WHERE sb_id="&request("sbid"))(0)&" 图片信息</b></div><br/>"
	dwt.out conn.Execute("SELECT sb_img FROM sbqt WHERE sb_id="&request("sbid"))(0)
  case "addsb"
      call addsb'选择分类后添加设备页面
  case "saveaddsb"
      call saveaddsb'设备添加保存
  case "edit"
      if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"'编辑子分类
      call saveedit'编辑保存子分类
  case "del"
        if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
			
	    Dwt.savesl "设备管理-"&zclass(conn.Execute("SELECT sb_dclass FROM sbqt WHERE sb_id="&request("id"))(0)),"删除",conn.Execute("SELECT sb_wh FROM sbqt WHERE sb_id="&request("id"))(0)

			Set Rs = Server.CreateObject("ADODB.Recordset")
			Sql = "Delete From sbqt Where sb_id="&request("id")
			Conn.execute(Sql)
			Dwt.out "<Script Language=Javascript>history.back()</Script>"
			set rs=nothing
			set conn=nothing
		end if 
  case ""
      if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	  	 

sub addsb()
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.sb_sscj.value==''){" & vbCrLf
Dwt.out "      alert('请选择所属车间！');" & vbCrLf
Dwt.out "   document.form.sb_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_ssgh.value==0){" & vbCrLf
Dwt.out "      alert('请选择所属装置！');" & vbCrLf
Dwt.out "   document.form.sb_ssgh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_wh.value==''){" & vbCrLf
Dwt.out "      alert('请添写位号！');" & vbCrLf
Dwt.out "   document.form.sb_wh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf


Dwt.out " if(document.form.sb_sccj.value==''){" & vbCrLf
Dwt.out "      alert('请添写生产厂家！');" & vbCrLf
Dwt.out "   document.form.sb_sccj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out"<form method='post' action='sb_qtjc.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	Dwt.out"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.out"<tr class='title'>"& vbCrLf
	Dwt.out"<td height='22' colspan='2'><Div align=center><strong>新增 "&sb_classname&" 设备</strong></Div></tr>"& vbCrLf
	
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
if session("level")=0 then 



	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	 Dwt.out"<select name='sb_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
     Dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
	 sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
     set rscj=server.createobject("adodb.recordset")
     rscj.open sqlcj,conn,1,1
     do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
	
    Dwt.out "<select name='sb_ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>选择班组分类</option>" & vbCrLf
    Dwt.out "</select>"  & vbCrLf
    Dwt.out "</select>  "  & vbCrLf
    Dwt.out "<select name='sb_ssgh' size='1' >" & vbCrLf
    Dwt.out "<option  selected>选择装置分类</option>" & vbCrLf
    Dwt.out "</select>"  & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
	
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form.sb_sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	
    Dwt.out "var groupgh=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "groupgh[i]=new Array()" & vbCrLf
    Dwt.out "groupgh[0][0]=new Option(""选择装置分类"","" "");" & vbCrLf
	
	
	
	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0
	 jj=0
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""无班组"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

	 sqlgh="SELECT * from ghname where sscj="&rscj("levelid") 
        set rsgh=server.createobject("adodb.recordset")
        rsgh.open sqlgh,conn,1,1
        if rsgh.eof and rsgh.bof then
		   Dwt.out "groupgh["&rscj("levelid")&"][0]=new Option(""无装置"",""0"");" & vbCrLf
		else
		do while not rsgh.eof
		   Dwt.out"groupgh["&rsgh("sscj")&"]["&jj&"]=new Option("""&rsgh("gh_name")&""","""&rsgh("ghid")&""");" & vbCrLf
		  jj=jj+1
		   rsgh.movenext
	    loop
	end if
		rsgh.close
	    set rsgh=nothing
		
		
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out "var temp=document.form.sb_ssbz" & vbCrLf
    Dwt.out "var temp2=document.form.sb_ssgh" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
	
    Dwt.out "for (i=0;i<groupgh[x].length;i++){" & vbCrLf
    Dwt.out "temp2.options[i]=new Option(groupgh[x][i].text,groupgh[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp2.options[0].selected=true" & vbCrLf
	
    Dwt.out "}//</script" & vbCrLf
		
    Dwt.out "</td> </tr> "  & vbCrLf
	
  else 	 
   Dwt.out"<input name='sb_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   Dwt.out"<input name='sb_sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   
	 Dwt.out"<select name='sb_ssbz' size='1'>"& vbCrLf
	 sqlbz="SELECT * from bzname where sscj="&session("levelclass")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
		do while not rsbz.eof
       	Dwt.out"<option value='"&rsbz("id")&"'>"&rsbz("bzname")&"</option>"& vbCrLf
	
		   rsbz.movenext
	    loop
		rsbz.close
	    set rsbz=nothing
    Dwt.out"</select>"  	 & vbCrLf
   
   
   
   if session("levelclass")=4 then 
      sql="SELECT * from ghname "
   else
      sql="SELECT * from ghname where sscj="&session("levelclass")
   end if 
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   Dwt.out"<select name='sb_ssgh' size='1'>"
   
   if rs.eof and rs.bof then 
   	  Dwt.out"<option value='0'>未添加装置</option>"
   else   
	  'Dwt.out"<option value='0'>车间</option>"
      do while not rs.eof
	     Dwt.out"<option value='"&rs("ghid")&"'>"&rs("gh_name")&"</option>"
	  rs.movenext
      loop
	  end if 
	 Dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
    Dwt.out"</td></tr>  "  	 

	
	if zclassor(sb_classid) then
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>类型： </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass sb_classid,0
		Dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>位号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_wh' type='text' ></td></tr>"& vbCrLf

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>安装位置： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_install_location' type='text' ></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>安全类型： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_security' type='text' ></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>精度等级： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_jddj' type='text' ></td></tr>"& vbCrLf

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>设备特性： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out" <label><input type='checkbox' name='sb_iszj'/>是否周检 </label>"

	Dwt.out" <label><input type='checkbox' name='sb_control_alarm'/>控制器是否报警 </label>"
	Dwt.out" <label><input type='checkbox' name='sb_alarm_itself'/>本体是否报警 </label>"
	Dwt.out" <label><input type='checkbox' name='sb_trend_record'/>趋势记录功能 </label>"
	Dwt.out" <label><input type='checkbox' name='sb_alarm_record'/>报警记录功能 </label>"

	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>鉴定周期：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict ("sb_test_period","鉴定周期",onnumb)
	dwt.out "</td></tr>"
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检查周期：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict ("sb_jczq","鉴定周期",onnumb)
	dwt.out "</td></tr>"
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>管理方式：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict ("sb_glfs","管理方式",onnumb)
	dwt.out "</td></tr>"
	
	Dwt.out "</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>完好： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_whqk' size='1' >"
	Dwt.out"<option value='0'"
	
	Dwt.out">请选择完好情况</option>"
	Dwt.out"<option value='1'>完好</option>"
	Dwt.out"<option value='2'>不完好</option>"
	Dwt.out"</select></td></tr>"
	if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&sb_classid)(0) then 
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>准确： </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_zqqk' size='1' >"
		Dwt.out"<option value='0'>请选择准确情况</option>"
		Dwt.out"<option value='1'>最大最小</option>"
		Dwt.out"<option value='2'>中间</option>"
		Dwt.out"<option value='3'>>95%</option>"
		Dwt.out"</select></td></tr>"
   end if 
	
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>投运： </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_tyqk' size='1' >"
		Dwt.out"<option value='0'>请选择投运情况</option>"
		Dwt.out"<option value='1'>投运</option>"
		Dwt.out"<option value='2'>原因未投运</option>"
		Dwt.out"<option value='3'>工艺原因未投运</option>"
		Dwt.out"</select></td></tr>"
   
   
   	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>分级： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_fj' size='1' >"
	Dwt.out"<option value='0'>请选择分级</option>"
	Dwt.out"<option value='1'>★</option>"
	Dwt.out"<option value='2'>★★</option>"
	Dwt.out"<option value='3'>★★★</option>"
	Dwt.out"</select></td></tr>"



	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>型号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_ggxh' type='text'>  <span class='tips'>输入空格显示所有已存在数据</span></td></tr>"& vbCrLf
	Dwt.out "<script>"
    '自动完成后面的内容为选中状态
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_ggxh"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_ggxh&btext=sbqt&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	
	
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		'Dwt.out "<p align=""center"">暂无内容</p>" 
	else
		do while not rsbody1.eof
			'字段名
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'字段在页面中显示的名称
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
		rsbody1.movenext
		loop
		sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
		sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  '去除最右边逗号
		sb_tablename=split(sb_tablename,",")
		sb_tablebody=split(sb_tablebody,",")
		for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
			dim sbtablename
			sbtablename=sb_tablename(sb_tablei)
			Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"： </strong></td>"   & vbCrLf   
			Dwt.out"<td width='60%' class='tdbg'><input name='"&sbtablename&"' type='text'></td></tr>"& vbCrLf
		next
	end if 
	set rsbody1=nothing	

	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>生产厂家： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_sccj' type='text'>  <span class='red'>*</span><span class='tips'>输入空格显示所有已存在数据</span></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>产品编号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bh' type='text' ></td></tr>"& vbCrLf
   	Dwt.out "<script>"
    '自动完成后面的内容为选中状态
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_sccj"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_sccj&btext=sbqt&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"

   
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>启用时间： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='88%' class='tdbg'>"
   Dwt.out"<input name='sb_qydate' type='text'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
    Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>备注： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bz' type='text'></td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>图片： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_img' id='sb_img' type='hidden' >"& vbCrLf
    Dwt.out "<iframe src='neweditor/editor.htm?id=sb_img&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
    dwt.out "</td></tr>"
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveaddsb'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>     <input  type='submit' name='Submit' value=' 添   加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

end sub

sub saveaddsb()
dim temp2

'新增保存
	sb_classid=request("sbclassid")
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sbqt"
	rsadd.open sqladd,conn,1,3
	rsadd.addnew
    on error resume next
    rsadd("sb_dclass")=ReplaceBadChar(Trim(Request("sbclassid")))
	rsadd("sb_sscj")=ReplaceBadChar(Trim(Request("sb_sscj")))
	rsadd("sb_ssbz")=ReplaceBadChar(Trim(Request("sb_ssbz")))
	rsadd("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsadd("sb_dclass")) then 	rsadd("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '判断是否有子分类,再保存
	rsadd("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	
	rsadd("sb_install_location")=ReplaceBadChar(Trim(Request("sb_install_location")))
	rsadd("sb_security")=ReplaceBadChar(Trim(Request("sb_security")))
	rsadd("sb_test_period")=ReplaceBadChar(Trim(Request("sb_test_period")))
	rsadd("sb_jczq")=ReplaceBadChar(Trim(Request("sb_jczq")))
	rsadd("sb_jddj")=ReplaceBadChar(Trim(Request("sb_jddj")))
	rsadd("sb_glfs")=ReplaceBadChar(Trim(Request("sb_glfs")))

    sb_control_alarm=request("sb_control_alarm")
	if sb_control_alarm="on" then
	  sb_control_alarm=true
	else
	  sb_control_alarm=false
	end if  
	rsadd("sb_control_alarm")=sb_control_alarm
   
    sb_alarm_itself=request("sb_alarm_itself")
	if sb_alarm_itself="on" then
	  sb_alarm_itself=true
	else
	  sb_alarm_itself=false
	end if  
	rsadd("sb_alarm_itself")=sb_alarm_itself

    sb_trend_record=request("sb_trend_record")
	if sb_trend_record="on" then
	  sb_trend_record=true
	else
	  sb_trend_record=false
	end if  
	rsadd("sb_trend_record")=sb_trend_record

    sb_alarm_record=request("sb_alarm_record")
	if sb_alarm_record="on" then
	  sb_alarm_record=true
	else
	  sb_alarm_record=false
	end if  
	rsadd("sb_alarm_record")=sb_alarm_record

	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsadd("sb_iszj")=sb_iszj
	
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsadd("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsadd("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsadd("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsadd("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
	rsadd("sb_bh")=ReplaceBadChar(Trim(request("sb_bh")))
	rsadd("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	
    if isnull(rsadd("sb_sczjdate")) then  rsadd("sb_sczjdate")=ReplaceBadChar(Trim(request("sb_qydate")))
    if isnull(rsadd("sb_scjcdate")) then  rsadd("sb_scjcdate")=ReplaceBadChar(Trim(request("sb_qydate")))
	rsadd("sb_img")=Trim(request("sb_img"))
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		'Dwt.out "<p align=""center"">暂无内容</p>" 
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
	rsadd("sb_update")=now()
	rsadd.update
	rsadd.close
	  Dwt.savesl "设备管理-"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&ReplaceBadChar(Trim(Request("sbclassid"))))(0),"添加",ReplaceBadChar(Trim(Request("sb_wh")))
 	
	
	Dwt.out"<Script Language=Javascript>location.href='sb_qtjc.asp?sbclassid="&sb_classid&"'</Script>"

end sub


sub saveedit()
  dim temp1
'编辑保存
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sbqt where sb_ID="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,conn,1,3
	on error resume next

	rsedit("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsedit("sb_dclass")) then 	rsedit("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '判断是否有子分类,再保存
	rsedit("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	
	rsedit("sb_install_location")=ReplaceBadChar(Trim(Request("sb_install_location")))
	rsedit("sb_security")=ReplaceBadChar(Trim(Request("sb_security")))
	rsedit("sb_test_period")=ReplaceBadChar(Trim(Request("sb_test_period")))
	rsedit("sb_jczq")=ReplaceBadChar(Trim(Request("sb_jczq")))
	rsedit("sb_jddj")=ReplaceBadChar(Trim(Request("sb_jddj")))
	rsedit("sb_glfs")=ReplaceBadChar(Trim(Request("sb_glfs")))

    sb_control_alarm=request("sb_control_alarm")
	if sb_control_alarm="on" then
	  sb_control_alarm=true
	else
	  sb_control_alarm=false
	end if  
	rsedit("sb_control_alarm")=sb_control_alarm
   
    sb_alarm_itself=request("sb_alarm_itself")
	if sb_alarm_itself="on" then
	  sb_alarm_itself=true
	else
	  sb_alarm_itself=false
	end if  
	rsedit("sb_alarm_itself")=sb_alarm_itself

    sb_trend_record=request("sb_trend_record")
	if sb_trend_record="on" then
	  sb_trend_record=true
	else
	  sb_trend_record=false
	end if  
	rsedit("sb_trend_record")=sb_trend_record
	
	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsedit("sb_iszj")=sb_iszj

    sb_alarm_record=request("sb_alarm_record")
	if sb_alarm_record="on" then
	  sb_alarm_record=true
	else
	  sb_alarm_record=false
	end if  
	rsedit("sb_alarm_record")=sb_alarm_record
	
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsedit("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsedit("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsedit("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsedit("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
	rsedit("sb_bh")=ReplaceBadChar(Trim(request("sb_bh")))
    rsedit("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	
    if isnull(rsedit("sb_sczjdate")) then  rsedit("sb_sczjdate")=ReplaceBadChar(Trim(request("sb_qydate")))
	if isnull(rsedit("sb_scjcdate")) then  rsedit("sb_scjcdate")=ReplaceBadChar(Trim(request("sb_qydate")))

	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		Dwt.out "<p align=""center"">暂无内容</p>" 
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
	
	rsedit("sb_img")=Trim(request("sb_img"))
	rsedit("sb_sccj")=ReplaceBadChar(Trim(request("sb_sccj")))
	rsedit("sb_bz")=ReplaceBadChar(Trim(request("sb_bz")))
	rsedit("sb_update")=now()
	rsedit.update
	rsedit.close
	  Dwt.savesl "设备管理-"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&ReplaceBadChar(Trim(Request("sbclassid"))))(0),"编辑",ReplaceBadChar(Trim(Request("sb_wh")))
	Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub edit()
	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.sb_sscj.value==''){" & vbCrLf
Dwt.out "      alert('请选择所属车间！');" & vbCrLf
Dwt.out "   document.form.sb_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_ssgh.value==0){" & vbCrLf
Dwt.out "      alert('请选择所属装置！');" & vbCrLf
Dwt.out "   document.form.sb_ssgh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_wh.value==''){" & vbCrLf
Dwt.out "      alert('请添写位号！');" & vbCrLf
Dwt.out "   document.form.sb_wh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf


Dwt.out " if(document.form.sb_sccj.value==''){" & vbCrLf
Dwt.out "      alert('请添写生产厂家！');" & vbCrLf
Dwt.out "   document.form.sb_sccj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
	sb_id=ReplaceBadChar(Trim(request("id")))

	sqledit="SELECT * from sbqt where sb_id="&sb_id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,conn,1,1
	Dwt.out"<form method='post' action='sb_qtjc.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	Dwt.out"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.out"<tr class='title'>"& vbCrLf
	Dwt.out"<td height='22' colspan='2'><Div align=center><strong>"&sb_classname&"编辑 "
	Dwt.out"位号:"& vbCrLf
	Dwt.out rsedit("sb_wh")&"</strong></Div></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
if session("level")=0 then 



	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	 sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
     set rscj=server.createobject("adodb.recordset")
     rscj.open sqlcj,conn,1,1
	 Dwt.out"<select name='sb_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
     Dwt.out"<option  selected >选择所属车间</option>"& vbCrLf
	 sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
     set rscj=server.createobject("adodb.recordset")
     rscj.open sqlcj,conn,1,1
     do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
	
    Dwt.out "<select name='sb_ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>选择班组分类</option>" & vbCrLf
    Dwt.out "</select>"  & vbCrLf
    Dwt.out "</select>  "  & vbCrLf
    Dwt.out "<select name='sb_ssgh' size='1' >" & vbCrLf
    Dwt.out "<option  selected>选择装置分类</option>" & vbCrLf
    Dwt.out "</select>"  & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
	
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form.sb_sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	
    Dwt.out "var groupgh=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "groupgh[i]=new Array()" & vbCrLf
    Dwt.out "groupgh[0][0]=new Option(""选择装置分类"","" "");" & vbCrLf
	
	
	
	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0
	 jj=0
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""无班组"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

	 sqlgh="SELECT * from ghname where sscj="&rscj("levelid") 
        set rsgh=server.createobject("adodb.recordset")
        rsgh.open sqlgh,conn,1,1
        if rsgh.eof and rsgh.bof then
		   Dwt.out "groupgh["&rscj("levelid")&"][0]=new Option(""无装置"",""0"");" & vbCrLf
		else
		do while not rsgh.eof
		   Dwt.out"groupgh["&rsgh("sscj")&"]["&jj&"]=new Option("""&rsgh("gh_name")&""","""&rsgh("ghid")&""");" & vbCrLf
		  jj=jj+1
		   rsgh.movenext
	    loop
	end if
		rsgh.close
	    set rsgh=nothing
		
		
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out "var temp=document.form.sb_ssbz" & vbCrLf
    Dwt.out "var temp2=document.form.sb_ssgh" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
	
    Dwt.out "for (i=0;i<groupgh[x].length;i++){" & vbCrLf
    Dwt.out "temp2.options[i]=new Option(groupgh[x][i].text,groupgh[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp2.options[0].selected=true" & vbCrLf
	
    Dwt.out "}//</script" & vbCrLf
		
    Dwt.out "</td> </tr> "  & vbCrLf
	
  else 	
	Dwt.out"<input name='sb_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sb_sscj"))&"'>"& vbCrLf
    Dwt.out"<input name='sb_sscj' type='hidden' value="&rsedit("sb_sscj")&">"& vbCrLf

	sqlbz="SELECT * from bzname where sscj="&rsedit("sb_sscj")
   	set rsbz=server.createobject("adodb.recordset")
   	rsbz.open sqlbz,conn,1,1
   	Dwt.Out"<select name='sb_ssbz' size='1'>"
   	if rsbz.eof and rsbz.bof then 
   		  Dwt.Out"<option value='0'>未添加班组</option>"& vbCrLf
   	else   
		  'Dwt.Out"<option value='0'>车间</option>"
   	   do while not rsbz.eof
		     Dwt.Out"<option value='"&rsbz("id")&"'"
			 if rsedit("sb_ssbz")=rsbz("id") then Dwt.Out " selected"
			 Dwt.Out">"&rsbz("bzname")&"</option>"& vbCrLf
		  rsbz.movenext
   	   loop
	end if 
	 Dwt.Out"</select>" & vbCrLf
 	 rsbz.close
 	 set rsbz=nothing
	Dwt.out"</td></tr>"

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>所属装置： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_ssgh' size='1' >"
	call formgh (rsedit("sb_ssgh"),rsedit("sb_sscj"))
	Dwt.out"</select>"
	end if
	Dwt.out"</td></tr>"
	
	if zclassor(rsedit("sb_dclass")) then
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>类型： </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass sb_classid,rsedit("sb_zclass") 
		Dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>位号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_wh' type='text' value='"&rsedit("sb_wh")&"'></td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>安装位置： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_install_location' type='text' value='"&rsedit("sb_install_location")&"' ></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>安全类型： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_security' type='text' value='"&rsedit("sb_security")&"' ></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>精度等级： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_jddj' type='text' value='"&rsedit("sb_jddj")&"' ></td></tr>"& vbCrLf
	
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>设备特性： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out" <label><input type='checkbox' name='sb_iszj' "
	if rsedit("sb_iszj") then Dwt.out "checked='checked'"
	Dwt.out" />是否周检 </label>"

	Dwt.out" <label><input type='checkbox' name='sb_control_alarm' "
	if rsedit("sb_control_alarm") then Dwt.out "checked='checked'"
	Dwt.out" />控制器是否报警 </label>"
	
	Dwt.out" <label><input type='checkbox' name='sb_alarm_itself' "
	if rsedit("sb_alarm_itself") then Dwt.out "checked='checked'"
	Dwt.out" />本体是否报警 </label>"

	Dwt.out" <label><input type='checkbox' name='sb_trend_record' "
	if rsedit("sb_trend_record") then Dwt.out "checked='checked'"
	Dwt.out" />趋势记录功能 </label>"

	Dwt.out" <label><input type='checkbox' name='sb_alarm_record' "
	if rsedit("sb_alarm_record") then Dwt.out "checked='checked'"
	Dwt.out" />报警记录功能 </label>"
	
	Dwt.out "</td></tr>"& vbCrLf
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>鉴定周期：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict2 ("sb_test_period","鉴定周期",onnumb,rsedit("sb_test_period"))
	dwt.out "</td></tr>"
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>检查周期：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
	if isnull(rsedit("sb_jczq")) then
    dwt.out outdatadict ("sb_jczq","鉴定周期",onnumb)
	else
    dwt.out outdatadict2 ("sb_jczq","鉴定周期",onnumb,rsedit("sb_jczq"))
    end if
	dwt.out "</td></tr>"
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>管理方式：</strong></td>"	
	Dwt.Out"<td width='80%' class='tdbg'>"	
    dwt.out outdatadict2 ("sb_glfs","管理方式",onnumb,rsedit("sb_glfs"))
	dwt.out "</td></tr>"


		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>完好： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_whqk' size='1' >"
	Dwt.out"<option value='0'"
	
	if rsedit("sb_whqk")=0 then Dwt.out" selected" 
	Dwt.out">请选择完好情况</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_whqk")=1 then Dwt.out"selected"
	Dwt.out">完好</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_whqk")=2 then Dwt.out"selected"
	Dwt.out" >不完好</option>"
	Dwt.out"</select></td></tr>"
	
	if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&sb_classid)(0) then 
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>准确： </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_zqqk' size='1' >"
		Dwt.out"<option value='0'"
		if rsedit("sb_zqqk")=0 then Dwt.out" selected" 
		Dwt.out">请选择准确情况</option>"
		Dwt.out"<option value='1' "
		if rsedit("sb_zqqk")=1 then Dwt.out"selected"
		Dwt.out">最大最小</option>"
		Dwt.out"<option value='2'"
		if rsedit("sb_zqqk")=2 then Dwt.out"selected"
		Dwt.out" >中间</option>"
		Dwt.out"<option value='3'"
		if rsedit("sb_zqqk")=3 then Dwt.out"selected"
		Dwt.out" >>95%</option>"
		Dwt.out"</select></td></tr>"
    end if 

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>投运： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_tyqk' size='1' >"
	Dwt.out"<option value='0'"
	if rsedit("sb_tyqk")=0 then Dwt.out" selected" 
	Dwt.out">请选择投运情况</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_tyqk")=1 then Dwt.out"selected"
	Dwt.out">投运</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_tyqk")=2 then Dwt.out"selected"
	Dwt.out" >原因未投运</option>"
	Dwt.out"<option value='3' "
	if rsedit("sb_tyqk")=3 then Dwt.out"selected"
	Dwt.out">工艺原因未投运</option>"
	Dwt.out"</select></td></tr>"

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>分级： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_fj' size='1' >"
	Dwt.out"<option value='0'"
	if rsedit("sb_fj")=0 then Dwt.out" selected" 
	Dwt.out">请选择分级</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_fj")=1 then Dwt.out"selected"
	Dwt.out">★</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_fj")=2 then Dwt.out"selected"
	Dwt.out" >★★</option>"
	Dwt.out"<option value='3' "
	if rsedit("sb_fj")=3 then Dwt.out"selected"
	Dwt.out">★★★</option>"
	Dwt.out"</select></td></tr>"
	
	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>型号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_ggxh' type='text' value='"&rsedit("sb_ggxh")&"'>  <span class='tips'>输入空格显示所有已存在数据</span></td></tr>"& vbCrLf
	Dwt.out "<script>"
    '自动完成后面的内容为选中状态
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_ggxh"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_ggxh&btext=sbqt&orderby=sb_ggxh&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	dim sb_tablename,sb_tablebody,sb_table
			'取字段的名称
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		 
	else
		do while not rsbody1.eof
			'字段名
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'字段在页面中显示的名称
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
			
		rsbody1.movenext
		loop
sb_tablename=left(sb_tablename,len(sb_tablename)-1)  '去除最右边逗号
	sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  '去除最右边逗号
	sb_tablename=split(sb_tablename,",")
	sb_tablebody=split(sb_tablebody,",")


	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
		
			Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"： </strong></td>"   & vbCrLf   
			Dwt.out"<td width='60%' class='tdbg'><input name='"&sbtablename&"' type='text' value='"&rsedit(sbtablename)&"'></td></tr>"& vbCrLf
	next
	end if 
	set rsbody1=nothing	

	
	

	

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>生产厂家： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_sccj' type='text' value='"&rsedit("sb_sccj")&"'>  <span class='red'>*</span><span class='tips'>输入空格显示所有已存在数据</span></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>产品编号： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bh' type='text' value='"&rsedit("sb_bh")&"'>  <span class='tips'>输入空格显示所有已存在数据</span></td></tr>"& vbCrLf
   	Dwt.out "<script>"
    '自动完成后面的内容为选中状态
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_sccj"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_sccj&btext=sbqt&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	
	
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>启用时间： </strong></td>"   & vbCrLf   
   Dwt.out"<td width='88%' class='tdbg'>"
   Dwt.out"<input name='sb_qydate' type='text' onClick='new Calendar(0).show(this)' readOnly  value="&rsedit("sb_qydate")&">"
   Dwt.out"</td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>备注： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bz' type='text' value='"&rsedit("sb_bz")&"'></td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>图片： </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input type='hidden' name='sb_img' id='sb_img' value='"&rsedit("sb_img")&"'>"& vbCrLf
    Dwt.out "<iframe src='neweditor/editor.htm?id=sb_img&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
	dwt.out "</td></tr>"& vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>   <input name='id' type='hidden'  value='"&Trim(Request("id"))&"'> <input  type='submit' name='Submit' value=' 保 存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	
	Dwt.out"</table></form>"
  rsedit.close
  set rsedit=nothing
  conn.close
  set conn=nothing

end sub


sub main()
 url= GetUrl
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function isDel(id){" & vbCrLf
Dwt.out "		if(confirm('您确定要删除此内容吗？')){" & vbCrLf
Dwt.out "			location.href='sb_qtjc.asp?action=del&id='+id;" & vbCrLf
Dwt.out "		}" & vbCrLf
Dwt.out "	}" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf

	sqlbody="SELECT * from sbqt where sb_isdel=false and sb_dclass="&sb_classid

	if sscjid<>"" then sqlbody=sqlbody&" and sb_sscj="&sscjid
	if ssghid<>"" then sqlbody=sqlbody&" and sb_ssgh="&ssghid
	if keys<>"" then sqlbody=sqlbody&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlbody=sqlbody&" and sb_zclass="&request("sbzclassid")
	if request("update")<>"" then 
    	sqlbody=sqlbody&" order by sb_update desc"
	else
    	sqlbody=sqlbody&" order by sb_sscj aSC,sb_ssgh asc,sb_wh asc"
	end if 

	set rsbody=server.createobject("adodb.recordset")
	rsbody.open sqlbody,conn,1,1

	if request("sscj")<>"" then title=sscjh(sscjid)&"－" 
	if request("ssgh")<>"" then title=gh(ssghid) &"－"
	if request("keyword")<>"" then title=" '"&keys&" '"&"－"
    title="－－"&title&sb_classname
	if request("sbzclassid")<>"" then title=title&"－"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&request("sbzclassid"))(0)
	
	
	Dwt.out "<Div style='left:6px;'>"
	Dwt.out "     <Div class='x-layout-panel-hd x-layout-title-center'>"
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>设备档案"&title&"</span>"
	Dwt.out "     </Div>"

        sqlcj="SELECT distinct sb_sscj from sbqt where  sb_isdel=false and sb_dclass="&sb_classid
		
		   sqlcj=sqlcj&" order by sb_sscj asc"
	   set rscj=server.createobject("adodb.recordset")
               rscj.open sqlcj,conn,1,1
               do while not rscj.eof
	       sscji=cint(rscj("sb_sscj"))
           ' for sscji=1 to 5 
	  sql="SELECT count(sb_id) FROM sbqt WHERE sb_dclass="&sb_classid&" and sb_sscj="&sscji
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	  sb_numb=sb_numb&sscjh_d(sscji)&":"&"<font color='#006600'>"&conn.Execute(sql)(0)&"</font>&nbsp;&nbsp;&nbsp;&nbsp;"
	   ' next
              rscj.movenext
	      loop
	      rscj.close
	      set rscj=nothing

	sql="SELECT count(sb_id) FROM sbqt WHERE sb_dclass="&sb_classid
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	totall= "<font color='#006600'>"&conn.Execute(sql)(0)&"</font>" 
	'Dwt.out "<Div class='pre'> <strong>维一："&v1&"</strong>   <strong>维二："&v2&"</strong>     <strong>维三："&v3&"</strong>     <strong>维四："&v4&"</strong>     <strong>综合："&zh&"</strong>     <strong>合计："&totall&"</strong></Div>"
	Dwt.out "<Div class='pre'>"&sb_numb&"合计:"&totall&"<br/>位号前加<span style=""color:#0033ff"">★</span>表示最近更新过&nbsp;&nbsp;完好<span style=""color:#006600"">★</span>不完好<span style=""color:#ff0000"">★</span> &nbsp;&nbsp;投运<span style=""color:#006600"">★</span>因未投运<span style=""color:#0000ff"">★</span>因工艺未投运<span style=""color:#ff0000"">★</span></Div>"& vbCrLf
	Dwt.out "<Div class='x-layout-container' style='top:0px;WIDTH: 1000px; POSITION: relative; HEIGHT: 543px'>"& vbCrLf
	Dwt.out "<Div class='x-layout-panel x-layout-panel-center' style='LEFT: 3px; WIDTH: 1000px; TOP: 3px; HEIGHT: 537px'>"& vbCrLf
	search	()
	
	if rsbody.eof and rsbody.bof then 
		message "<p align=""center"">未找到相关内容</p>" & vbCrLf
	else
	    Dwt.out "<SCRIPT src='js/grid.js' type=text/javascript></SCRIPT>"& vbCrLf
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
		
		Dwt.out "<SCRIPT language=JavaScript >"& vbCrLf
        'Dwt.out "// 栏位标题 ( 栏位名称 # 栏位宽度 # 资料对齐 )"
		Dwt.out "var DataTitles=new Array("& vbCrLf
		Dwt.out """序号#40#center"","& vbCrLf
		Dwt.out """位号#160#left"","& vbCrLf
		Dwt.out """车间#120#center""  ,"& vbCrLf
		Dwt.out """装置#90#center"","& vbCrLf
		
		Dwt.out """安装位置#90#center"","& vbCrLf
		
		if zclassor(rsbody("sb_dclass")) then Dwt.out """类型#80 #center"","& vbCrLf
		Dwt.out """完好#58 #center"","& vbCrLf
		
		'如果在分类里设定了显示"准确"才显示
		if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&rsbody("sb_dclass"))(0) then Dwt.out """准确#58 #center"","& vbCrLf
		
		Dwt.out """投运#58 #center"","& vbCrLf
		Dwt.out """分级#58 #center"","& vbCrLf
		Dwt.out """型号#150#left"","& vbCrLf
		
		Dwt.out """安全类型#50#left"","& vbCrLf

		'取字段的名称
		sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
		set rsbody1=server.createobject("adodb.recordset")
		rsbody1.open sqlbody1,conn,1,1
		if rsbody1.eof and rsbody1.bof then 
			'Dwt.out "<p align=""center"">暂无内容</p>" 
		else
			do while not rsbody1.eof
				Dwt.out """"&rsbody1("sbtable_body")&"#140 #center"","& vbCrLf
				rsbody1.movenext
			loop
		end if 
		set rsbody1=nothing	
		
		Dwt.out """控制器报警#80 #left"","& vbCrLf
		Dwt.out """本体报警#80 #left"","& vbCrLf
		Dwt.out """趋势记录#80 #left"","& vbCrLf
		Dwt.out """报警记录#80 #left"","& vbCrLf
		
		Dwt.out """生产厂家#80 #left"","& vbCrLf
		Dwt.out """产品编号#150#left"","& vbCrLf
		Dwt.out """启用时间#70 #center"","& vbCrLf
		Dwt.out """备注#100 #left"","& vbCrLf
		Dwt.out """选项#80 #center"")</SCRIPT>"
		Dwt.out "<SCRIPT language=JavaScript >"
		Dwt.out "var DataFields=new Array()"& vbCrLf
		i=0
		do while not rsbody.eof and rowcount>0
				xh_id=((page-1)*pgsz)+1+xh
				xh=xh+1
			Dwt.out "DataFields["&i&"] =new Array("
			Dwt.out "'"&xh_id&"',"
			
			Dwt.out "'"
			if now()-rsbody("sb_update")<7 then Dwt.out "<span style=""color:#0033ff"">★</span>"
			Dwt.out searchH(uCase(rsbody("sb_wh")),keys)&"',"
			Dwt.out "'<a href=sb_jxjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&sb_classid&">检</a>&nbsp;<a href=sb_ghjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&sb_classid&">换</a>&nbsp;"
	        if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,rsbody("sb_sscj")) then Dwt.out "<a href=""sb_qtjc.asp?action=edit&sbclassid="&sb_classid&"&id="&rsbody("sb_id")&""">编</a>&nbsp;"
			if conn.Execute("SELECT sb_img FROM sbqt WHERE sb_id="&rsbody("sb_id"))(0)<>"" then dwt.out "<a href=sb_qtjc.asp?action=img&sbid="&rsbody("sb_id")&"  target=""_blank"">图</a>&nbsp;"
			Dwt.out sscjh_d(rsbody("sb_sscj"))&"',"
			Dwt.out "'"&GH(rsbody("sb_ssGH"))&"',"
			
			Dwt.out "'"&rsbody("sb_install_location")&"',"
			
			if zclassor(rsbody("sb_dclass")) then 
			   if zclass(rsbody("sb_zclass"))="未编辑" then 
			    dwt.out  "'"&zclass(rsbody("sb_dclass"))&"',"
			   else
			    Dwt.out "'"&zclass(rsbody("sb_zclass"))&"'," 
			   end if 
			 end if   	
			'Dwt.out """"&xh_id&""","
			Dwt.out "'"&sb_whd(rsbody("sb_whqk"))&"',"
			if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&rsbody("sb_dclass"))(0) then Dwt.out "'"&sb_zqd(rsbody("sb_ZQqk"))&"',"
			Dwt.out "'"&sb_tyd(rsbody("sb_tyqk"))&"',"
			Dwt.out "'"&fj(rsbody("sb_fj"))&"',"
			Dwt.out "'"&rsbody("sb_ggxh")&"',"

			Dwt.out "'"&rsbody("sb_security")&"',"
		
			'取字段内容
			sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
			set rsbody1=server.createobject("adodb.recordset")
			rsbody1.open sqlbody1,conn,1,1
			if rsbody1.eof and rsbody1.bof then 
				'Dwt.out "<p align=""center"">暂无内容</p>" 
			else
				do while not rsbody1.eof
				  sbtable_name=rsbody1("sbtable_name")   '取得表的名称
				  Dwt.out "'"&rsbody(sbtable_name)&"',"
				  'message sbtable_name
				rsbody1.movenext
				loop
			end if 
			set rsbody1=nothing	
			
			if rsbody("sb_control_alarm") then 
			 Dwt.out "'是 ',"
			 else Dwt.out "'否',"
			 end if
			if rsbody("sb_alarm_itself") then 
			 Dwt.out "'是 ',"
			 else Dwt.out "'否',"
			 end if
			if rsbody("sb_trend_record") then 
			 Dwt.out "'是 ',"
			 else Dwt.out "'否',"
			 end if
			if rsbody("sb_alarm_record") then 
			 Dwt.out "'是 ',"
			 else Dwt.out "'否',"
			 end if 

			Dwt.out "'"&rsbody("sb_sccj")&"',"
			Dwt.out "'"&rsbody("sb_bh")&"',"
			Dwt.out "'"&rsbody("sb_qydate")&"',"
			Dwt.out "'"&rsbody("sb_bz")&"',"
			Dwt.out "'"
			call sbeditdel(rsbody("sb_id"),rsbody("sb_sscj"),"sb_qtjc.asp?action=edit&sbclassid="&sb_classid&"&id=")'检修、更换、编辑、删除
			Dwt.out "')"& vbCrLf
			
			RowCount=RowCount-1
					i=i+1
		rsbody.movenext
		loop
		Dwt.out "</script>"
	Dwt.out "<TABLE cellSpacing=0 cellPadding=0 border=0>"
	Dwt.out "  <TBODY>"
	Dwt.out "  <tr>"
	Dwt.out "    <TD valign='top' style='BORDER-RIGHT: white 2px inset; BORDER-TOP: white 2px inset; BORDER-LEFT: white 2px inset; BORDER-BOTTOM: white 2px inset; BACKGROUND-COLOR: scrollbar'>"
	Dwt.out "      <Div id=DataTable></Div></TD></tr></TBODY></TABLE>"
		call sbshowpage(page,url,total,record)
	Dwt.out "</Div></Div></Div>"
	end if
	rsbody.close
	set rsbody=nothing
	conn.close
	set conn=nothing

end sub
'	Dwt.out "程序执行用时：" & timer-starttime

Dwt.out "</body></html>"

'选项（编辑、删除）
sub sbeditdel(id,sscj,editurl)
'	if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,sscj) then 
'	
'	Dwt.out "<a href="""&editurl&id&""">编辑</a>&nbsp;"
'	end if 	
	if  displaypagelevelh(session("groupid"),3,session("pagelevelid")) and displaygrouplevelh(session("groupid"),1,sscj)  then
	 Dwt.out "<a href=""javascript:isDel("&id&");"">删除</a>"
	end if 
	Dwt.out "&nbsp;"
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
	'if isnull(dclassid) then dclassid=0
'	if dclassid=0 then 
		sqldclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_isput=true"
'	else
'		sqldclass="SELECT * from sbclass where sbclass_dclass<>0 and sbclass_id="&dclassid
'	end if 		
	set rsdclass=server.createobject("adodb.recordset")
	rsdclass.open sqldclass,conn,1,1
	if rsdclass.eof then 
		dclass="没有任何分类"
	else
		Dwt.out"<option value='0'"
		if dclassid=0 then Dwt.out " selected" 
			Dwt.out">请选择要添加设备的分类</option>"
		do while not rsdclass.eof
			Dwt.out"<option value='sb_qtjc.asp?action=addsb&sbclassid="&rsdclass("sbclass_id")&"'>"&rsdclass("sbclass_name")&"</option>"  & vbCrLf   
		rsdclass.movenext
		loop
	end if 
	rsdclass.close
	set rsdclass=nothing
end function


'子分类列表显示
function formzclass(dclassid,zclassid)
	dim sqlzclass,rszclass
	if isnull(zclassid) then zclassid=0
'	if zclassid=0 then 
		sqlzclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_zclass="&dclassid
'	else
		'sqlzclass="SELECT * from sbclass where sbclass_zclass<>0 and sbclass_id="&zclassid
'	end if 		
	set rszclass=server.createobject("adodb.recordset")
	rszclass.open sqlzclass,conn,1,1
	if rszclass.eof then 
		formzclass="未编辑"
	else
		Dwt.out"<option value='0'"
		if zclassid=0 then Dwt.out " selected" 
			Dwt.out">请选择类型</option>"
		do while not rszclass.eof
			Dwt.out"<option value='"&rszclass("sbclass_id")&"' "
			if zclassid=rszclass("sbclass_id") then Dwt.out "selected"
			Dwt.out">"&rszclass("sbclass_name")&"</option>"  & vbCrLf   
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
		if whnumb=1 then sb_whd="<span style=""color:#006600"">★</span>"  '完好绿
		if whnumb=2 then sb_whd="<span style=""color:#ff0000"">★</span> "	  '不完好红
	end if 
end Function 

'准确情况显示
Function sb_zqd(zqnumb)

	if isnull(zqnumb) or zqnumb=0 then 
	  sb_zqd="未编辑"
	else
		if zqnumb=3 then sb_zqd="★★★"'>95%
		if zqnumb=2 then sb_zqd="★★"		  '中间  
		if zqnumb=1 then sb_zqd="★"  '最大最小
	end if 
end Function 

'投运情况显示
Function sb_tyd(tynumb)

	if isnull(tynumb) or tynumb=0 then 
	  sb_tyd="未编辑"
	else
		if tynumb=1 then sb_tyd="<span style=""color:#006600"">★</span>"   '绿投运
		if tynumb=2 then sb_tyd="<span style=""color:#0000ff"">★</span>"   '蓝仪原因未投运
		if tynumb=3 then sb_tyd="<span style=""color:#ff0000"">★</span>"    '红工艺原因未投运
		'if zqnumb=4 then sb_zqd="<font color='#ff0000'>*</font>"    '红工艺原因未投运
	end if 
end Function 



sub search()
	dim rscj,sqlcj,sscjid
	Dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	
	Dwt.out "<form method='Get' name='SearchForm' action='sb_qtjc.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.out "<a href='sb_qtjc.asp?action=addsb&sbclassid="&sb_classid&"'>添加设备</a>"
	Dwt.out "&nbsp;&nbsp;<select   onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>显示顺序选择</option>" & vbCrLf
	Dwt.out "<option value='sb_qtjc.asp?update=update&sbclassid="&sb_classid&"'>按更新时间</option>"
	Dwt.out "     </select>	" & vbCrLf

	
	Dwt.out "  <input type='hidden' name='sbclassid' value='"&sb_classid&"'>" & vbCrLf
	if request("sbzclassid")<>"" then Dwt.out "<input type='hidden' name='sbzclassid' value='"&request("sbzclassid")&"'>" & vbCrLf

	Dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 Dwt.out "value='"&request("keyword")&"'"
    	Dwt.out ">" & vbCrLf
    else
	 Dwt.out "value='输入搜索的位号'"
	 	Dwt.out " onblur=""if(this.value==''){this.value='输入搜索的位号'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	Dwt.out "  <input type='submit' name='Submit'  value='搜索'>"
	



	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tosscj(this.options[this.selectedIndex].value);}"">" & vbCrLf
	Dwt.out "<option value=''>按车间跳转至…</option>" & vbCrLf
	sqlgh="SELECT distinct sb_sscj from sbqt where sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		cjid=cint(rsgh("sb_sscj"))


		sql="SELECT count(sb_id) FROM sbqt WHERE sb_sscj="&cjid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			'i=i+1
			Dwt.out"<option value='"&cjid&"'"
			if cint(request("sscj"))=cjid then Dwt.out" selected"
			sql="SELECT levelname FROM levelname WHERE levelid="&cjid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    cj_name="未知项"
			else
			    cj_name=rs("levelname")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&cj_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf

















'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'		set rscj=server.createobject("adodb.recordset")
'		rscj.open sqlcj,conn,1,1
'		do while not rscj.eof
'			Dwt.out"<option value='"&rscj("levelid")&"' "
'			if cint(request("sscj"))=rscj("levelid") then Dwt.out" selected"
'			Dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf	
'			rscj.movenext
'		loop
'		rscj.close
'		set rscj=nothing
'		Dwt.out "     </select>	" & vbCrLf



'Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tossgh(this.options[this.selectedIndex].value);}"">" & vbCrLf
'Dwt.out "	       <option value=''>按装置跳转至…</option>" & vbCrLf
'sscjid=session("levelclass")
'if sscjid=7 or sscjid=8 then
'sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
'else
'sqlgh="SELECT * from ghname where sscj="&sscjid&" ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
'end if
'    set rsgh=server.createobject("adodb.recordset")
'    rsgh.open sqlgh,conn,1,1
'    do while not rsgh.eof
'		sb_numb=Conn.Execute("SELECT count(sb_id) FROM sbqt WHERE sb_ssgh="&rsgh("ghid")&"and sb_dclass="&sb_classid)(0)
'		if sb_numb<>0 then 
'			i=i+1
'			Dwt.out"<option value='"&rsgh("ghid")&"'"
'			if cint(request("ssgh"))=rsgh("ghid") then Dwt.out" selected"
'			Dwt.out ">"&i&"&nbsp;&nbsp;"&rsgh("gh_name")&"("&sb_numb&")</option>"& vbCrLf
'	    end if 
'		rsgh.movenext
'	loop
'	rsgh.close
'	set rsgh=nothing
'	Dwt.out "     </select>	" & vbCrLf
	
	
	

	
	
Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tossgh(this.options[this.selectedIndex].value);}"">" & vbCrLf
Dwt.out "	       <option value=''>按装置跳转至…</option>" & vbCrLf



	sqlgh="SELECT distinct sb_ssgh,sb_sscj from sbqt where sb_isdel=false and sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		ghid=cint(rsgh("sb_ssgh"))


		sql="SELECT count(sb_id) FROM sbqt WHERE sb_isdel=false and  sb_ssgh="&ghid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			i=i+1
			Dwt.out"<option value='"&ghid&"'"
			if cint(request("ssgh"))=ghid and request("ssgh")<>"" then Dwt.out" selected"
			
			sql="SELECT gh_name FROM ghname WHERE ghid="&ghid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    gh_name="未知项"
			else
			    gh_name=rs("gh_name")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&i&"&nbsp;&nbsp;"&gh_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf


	Dwt.out "</form></Div></Div>" & vbCrLf
	
	
end sub

'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
'URL中带？的
'*******************************************
sub sbshowpage(page,url,total,record)
   Dwt.Out "<Div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      Dwt.Out "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      Dwt.Out ""
   end if 
   if RowCount = 0 and page <>Total then 
     Dwt.Out "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     Dwt.Out ""
   end if
   Dwt.Out"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  'if Total =1 then 
  '  Dwt.Out"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  'else
  ' Dwt.Out"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
 ' end if 
   if Total =1 then 
    Dwt.Out"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    Dwt.Out"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 Dwt.Out"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 Dwt.Out"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   Dwt.Out" </select>&nbsp;&nbsp;共"&record&"条内容"
   Dwt.Out "</Div>"
end sub

%>