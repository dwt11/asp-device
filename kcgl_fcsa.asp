<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->

<%
dim rs,sql

response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>库存台账管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "<SCRIPT language=javascript>" & vbCrLf
response.write "function checkfc(){" & vbCrLf
response.write "  if(document.form1.kcgl_numb.value==''){" & vbCrLf
response.write "      alert('数量未添写！');" & vbCrLf
response.write "  document.form1.kcgl_numb.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "  if(document.form1.kcgl_qxtxt.value==''){" & vbCrLf
response.write "      alert('出库对象未选择！');" & vbCrLf
response.write "  document.form1.kcgl_qxtxt.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf
response.write "    }" & vbCrLf

response.write "function checkadd(){" & vbCrLf
response.write "  if(document.form1.kcgl_sscj.value==''){" & vbCrLf
response.write "      alert('所属车间未选择！');" & vbCrLf
response.write "  document.form1.kcgl_sscj.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf

response.write "  if(document.form1.zclass.value==''){" & vbCrLf
response.write "      alert('子分类选择！');" & vbCrLf
response.write "  document.form1.zclass.focus();" & vbCrLf
response.write "      return false;" & vbCrLf
response.write "    }" & vbCrLf


response.write "    }" & vbCrLf


response.write "function checkamoney(){" & vbCrLf
response.write "if(document.getElementById(""checkamoney"").style.display==""none"")" & vbCrLf
response.write "		document.getElementById(""checkamoney"").style.display=""inline"";" & vbCrLf
		
response.write "	var szdmoney=document.getElementById(""kcgl_dmoney"").value;" & vbCrLf
response.write "	var sznumb=document.getElementById(""kcgl_numb"").value;" & vbCrLf
response.write "	if(szdmoney=="""")" & vbCrLf
response.write "	{	" & vbCrLf
response.write "		document.getElementById(""checkamoney"").innerHTML="" 正确输入单价能自动计算出金额!"";" & vbCrLf
response.write "		document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
response.write "		     return;}else" & vbCrLf

response.write "	      if(sznumb=="""")" & vbCrLf
response.write "	      {	" & vbCrLf
response.write "		      document.getElementById(""checkamoney"").innerHTML="" 正确输入数量能自动计算出金额!"";" & vbCrLf
response.write "		      document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
response.write "		     return;" & vbCrLf
response.write "	}" & vbCrLf

response.write "	var szamoney=document.getElementById(""kcgl_numb"").value*document.getElementById(""kcgl_dmoney"").value;" & vbCrLf

response.write "	document.getElementById(""checkamoney"").innerHTML=szamoney;" & vbCrLf
response.write "	document.getElementById(""checkamoney"").className=""ok"";" & vbCrLf
response.write "	return;" & vbCrLf

response.write "    }" & vbCrLf
response.write "</SCRIPT>" & vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write "<br><br>"
if Request("action")="fc" then call fc '主页面显示最新库存信息
if Request("action")="savefc" then call savefc   '保存出库内容(向现存表\入库表\出库表中写值)

'call ylb_search1(1)
if Request("action")="sr" then call sr '主页面显示最新库存信息
if Request("action")="savesr" then call savesr   '保存出库内容(向现存表\入库表\出库表中写值)

if Request("action")="add" then call add '入库添加
if Request("action")="saveadd" then call saveadd '入库添加保存

sub add()
dim rscj,sqlcj
   response.write"<form method='post' action='kcgl_fcsa.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>新物品入库添加</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
    response.write"<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	response.write"<select name='kcgl_sscj' size='1'>"
    response.write"<option >请选择所属车间</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select></td></tr>  "  	 
  else 	 
     response.write"<input name='kcgl_sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
      response.write"<input name='kcgl_sscj' type='hidden' value="&session("level")&"></td></tr>"& vbCrLf

 end if 

	 
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>分&nbsp;&nbsp;类：</strong></td>"
     response.write "<td><select name='dclass' size='1' onChange=""redirect(this.options.selectedIndex)"">" & vbCrLf
     response.write "<option  selected>选择父分类</option> "

     sql="SELECT * from class"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
    if rs.eof and rs.bof then 
          response.write"暂无分类"
      else
	  do while not rs.eof
          response.write"<option value='"&rs("id")&"'>"&rs("name")&"</option>"
	  rs.movenext
	loop
    end if 
    rs.close
    set rs=nothing
	response.write "</select>" & vbCrLf
    response.write "<select name='zclass' size='1' >" & vbCrLf
    response.write "<option  selected>选择子分类</option>" & vbCrLf
    response.write "</select></td></tr>" & vbCrLf
	
	
	
	response.write "<script>" & vbCrLf
response.write "<!--" & vbCrLf


response.write "var groups=document.form1.dclass.options.length" & vbCrLf
response.write "var group=new Array(groups)" & vbCrLf
response.write "for (i=0; i<groups; i++)" & vbCrLf
response.write "group[i]=new Array()" & vbCrLf
response.write"group[0][0]=new Option(""选择子分类"",""0"");" & vbCrLf
dim sqld,rsd,rsz,sqlz
sqld="SELECT * from class"
    set rsd=server.createobject("adodb.recordset")
    rsd.open sqld,connkc,1,1
    if rsd.eof and rsd.bof then 
          response.write"暂无分类"
      else
	  do while not rsd.eof
          sqlz="SELECT * from kcclass where class="&rsd("id")
         set rsz=server.createobject("adodb.recordset")
         rsz.open sqlz,connkc,1,1
         dim ia
		 ia=0
		 if rsz.eof and rsz.bof then 
            response.write"group["&rsd("id")&"]["&ia&"]=new Option(""无子分类"","""");" & vbCrLf
         else
		 
	        do while not rsz.eof
			        response.write"group["&rsd("id")&"]["&ia&"]=new Option("""&rsz("name")&""","""&rsz("id")&""");" & vbCrLf
	        
			ia=ia+1
			rsz.movenext
	        loop
         end if 
         rsz.close
	  rsd.movenext
	loop
    end if 
    rsd.close
    set rsd=nothing
response.write"var temp=document.form1.zclass" & vbCrLf
response.write"function redirect(x){" & vbCrLf
response.write"for (m=temp.options.length-1;m>0;m--)" & vbCrLf
response.write"temp.options[m]=null" & vbCrLf
response.write"for (i=0;i<group[x].length;i++){" & vbCrLf
response.write"temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
response.write"}" & vbCrLf
response.write"temp.options[0].selected=true" & vbCrLf
response.write"}" & vbCrLf
response.write"//-->" & vbCrLf
response.write"</script>" & vbCrLf

	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>名&nbsp;&nbsp;称：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_name' ></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>规格型号：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_xhgg'></td></tr> "
	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单位：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_dw'></td></tr>  "   
   
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单价：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_dmoney'  onBlur=""checkamoney()"" >元</td></tr>  "   
   
   	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数量：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_numb'  onBlur=""checkamoney()"" ></td></tr>  "   
   	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>金额：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><div id=""checkamoney"" style=""display:none"" class=""ok""></div>元</td></tr>  "   
	 	
		 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>入库日期：</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='kcgl_srdate' type='text' value="&now()&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this,kcgl_srdate, 'yyyy-mm-dd'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '保存到显存表中
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xc" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
      'sscj=request("kcgl_sscj")
		  'if sscj="" then sscj=7
       dim xcid
	  xcid=rsadd("id")
   	  rsadd("wpid")=xcid
	  rsadd("class")=request("zclass")
      'rsadd("lytxt")=request("kcgl_lytxt")
	  rsadd("sscj")=request("kcgl_sscj")
      on error resume next
      rsadd("name")=Trim(request("kcgl_name"))
      rsadd("xhgg")=request("kcgl_xhgg")
      rsadd("dw")=request("kcgl_dw")
      rsadd("dmoney")=request("kcgl_dmoney")
      rsadd("numb")=request("kcgl_numb")
      rsadd("amoney")=request("kcgl_dmoney")*request("kcgl_numb")
      rsadd("bz")=request("kcgl_bz")
	  rsadd("rcdate")=request("kcgl_srdate")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  
	  	  '保存到收入表中
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from sr" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
   	  rsadd("wpid")=xcid
rsadd("class")=request("zclass")
      rsadd("sscj")=request("kcgl_sscj")
      on error resume next
      rsadd("lytxt")=request("kcgl_lytxt")
	  rsadd("name")=Trim(request("kcgl_name"))
      rsadd("xhgg")=request("kcgl_xhgg")
      rsadd("dw")=request("kcgl_dw")
      rsadd("dmoney")=request("kcgl_dmoney")
      rsadd("numb")=request("kcgl_numb")
      rsadd("amoney")=request("kcgl_dmoney")*request("kcgl_numb")
      dim srdate
	  srdate=request("kcgl_srdate")
	  if srdate="" then srdate=year(now())&"-"&month(now())&"-"&day(now())
	  rsadd("sr_year")=year(srdate)
  	  rsadd("sr_month")=month(srdate)
  	  rsadd("sr_day")=day(srdate)
	  rsadd("bz")=request("kcgl_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='kcgl.asp';</Script>"
end sub

'入库
sub sr()
dim id 
dim rscj,sqlcj
dim classname
dim rssr,sqlsr
   id=ReplaceBadChar(Trim(request("id")))
   set rssr=server.createobject("adodb.recordset")
   sqlsr="select * from xc where id="&id
   rssr.open sqlsr,connkc,1,1
   
   response.write"<form method='post' action='kcgl_fcsa.asp' name='form1' onsubmit='javascript:return checkfc();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>库存台账---入库信息添写</strong></div></td>    </tr>"
	
 
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     response.write"<td width='88%' class='tdbg'><input type='text' value='"&sscjh(rssr("sscj"))&"' disabled='disabled' ></td></tr>"& vbCrLf
	response.write"<input name='kcgl_sscj' type='hidden' value="&rssr("sscj")&">"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>分&nbsp;&nbsp;类：</strong></td>"
	'sql="SELECT * from kcclass where id="&rssr("class")
    'set rs=server.createobject("adodb.recordset")
    'rs.open sql,connkc,1,1
    'if rs.eof and rs.bof then 
    '      classname="暂无分类"
    '  else
	'    classname=rs("name")
   ' end if 
   ' rs.close
   ' set rs=nothing
	response.write"<td><input value="&kcclass(rssr("class"))&"  disabled='disabled' ></td></tr>"
	response.write"<input name='kcgl_class' type='hidden' value="&rssr("class")&"></td></tr>"& vbCrLf
	 	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>来&nbsp;&nbsp;源：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rssr("lytxt")&">此项用于添写外购来源，如分厂-车间，车间-车间之间可不添写</td></tr> "
     response.write"<input  name='kcgl_lytxt'  type='hidden' value="&rssr("lytxt")&"></td></tr>"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>名&nbsp;&nbsp;称：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rssr("name")&"></td></tr> "
	 response.write"<input name='kcgl_name' type='hidden' value="&rssr("name")&">"
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>规格型号：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rssr("xhgg")&"></td></tr> "
	 response.write"<input name='kcgl_xhgg' type='hidden' value="&rssr("xhgg")&">"

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单位：</strong></td>"      
     response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rssr("dw")&"></td></tr>  "   
	 response.write"<input  name='kcgl_dw'  type='hidden' value="&rssr("dw")&">"
   
   
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单价：</strong></td>"      
     response.write"<td width='88%' class='tdbg'><input type='text'disabled='disabled' value="&rssr("dmoney")&">元</td></tr>  "   
 	 response.write"<input  name='kcgl_dmoney' type='hidden' value="&rssr("dmoney")&">"
  
   	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数量：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_numb'  onBlur=""checkamoney()""></td></tr>  "   
   	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>金额：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><div id=""checkamoney"" style=""display:none"" class=""ok""></div>元</td></tr>  "   
	 	
		 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>入库日期：</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='kcgl_srdate' type='text' value="&now()&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this,kcgl_srdate, 'yyyy-mm-dd'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='savesr'>  <input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 完 成 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
       rssr.close
       set rssr=nothing

end sub



sub savesr()    '保存出存信息
	  '保存到出库表中
	  dim rssr,sqlsr
	  dim sscj
	  dim rsadd,sqladd
	  	  '保存到收入表中
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from sr" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
   	  rsadd("xcid")=request("id")
      rsadd("class")=request("kcgl_class")
      rsadd("sscj")=request("kcgl_sscj")
      on error resume next
      rsadd("lytxt")=request("kcgl_lytxt")
	  rsadd("name")=Trim(request("kcgl_name"))
      rsadd("xhgg")=request("kcgl_xhgg")
      rsadd("dw")=request("kcgl_dw")
      rsadd("dmoney")=request("kcgl_dmoney")
      rsadd("numb")=request("kcgl_numb")
      rsadd("amoney")=request("kcgl_dmoney")*request("kcgl_numb")
      dim srdate
	  srdate=request("kcgl_srdate")
	  if srdate="" then srdate=year(now())&"-"&month(now())&"-"&day(now())
	  rsadd("sr_year")=year(srdate)
  	  rsadd("sr_month")=month(srdate)
  	  rsadd("sr_day")=day(srdate)
	  rsadd("bz")=request("kcgl_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing

'	  '编辑现存表中数量＿原数量＋表单提交过来的数量
	  dim rsedit,sqledit
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from xc where id="&request("id")
      rsedit.open sqledit,connkc,1,3
      rsedit("numb")=rsedit("numb")+request("kcgl_numb")
      rsedit("amoney")=rsedit("kcgl_dmoney")*rsedit("numb")
	  	  if srdate="" then srdate=year(now())&"-"&month(now())&"-"&day(now())
	  rsedit("rcdate")=srdate

	  rsedit.update
      rsedit.close
      set rsedit=nothing
	response.write"<Script Language=Javascript>history.go(-2)</Script>"
end sub


'出库
sub fc()
   dim id 
   dim rscj,sqlcj
   dim classname
   dim rsfc,sqlfc
   id=ReplaceBadChar(Trim(request("id")))
   set rsfc=server.createobject("adodb.recordset")
   sqlfc="select * from xc where id="&id
   rsfc.open sqlfc,connkc,1,1
   response.write"<form method='post' action='kcgl_fcsa.asp' name='form1' onsubmit='javascript:return checkfc();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>库存台账---出库信息添写</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>出库对象： </strong></td>"      
   response.write"<td width='88%' class='tdbg' id=txt>"
   response.write"<select name='kcgl_qxtxt' size='1'>"
   response.write"<option >请选择出库对象</option>"
   sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
   set rscj=server.createobject("adodb.recordset")
   rscj.open sqlcj,conn,1,1
   do while not rscj.eof
       	'出库对象下拉列表中不显示已属于的车间
		if rscj("levelid")=rsfc("sscj") then 
		  response.write""
		else
		   response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    end if 
		rscj.movenext
	loop
	response.write"<option value=1000>现场使用</option>"& vbCrLf
	rscj.close
	set rscj=nothing
   response.write"</select></td></tr>  "  	 
 
     response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     response.write"<td width='88%' class='tdbg'><input type='text' value='"&sscjh(rsfc("sscj"))&"' disabled='disabled' ></td></tr>"& vbCrLf
	response.write"<input name='kcgl_sscj' type='hidden' value="&rsfc("sscj")&">"& vbCrLf

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>分&nbsp;&nbsp;类：</strong></td>"
	response.write"<td><input value="&kcclass(rsfc("class"))&"  disabled='disabled' ></td></tr>"
	response.write"<input name='kcgl_class' type='hidden' value="&rsfc("class")&"></td></tr>"& vbCrLf
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>名&nbsp;&nbsp;称：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rsfc("name")&"></td></tr> "
	 response.write"<input name='kcgl_name' type='hidden' value="&rsfc("name")&">"
	
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>规格型号：</strong></td> "
	 response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rsfc("xhgg")&"></td></tr> "
	 response.write"<input name='kcgl_xhgg' type='hidden' value="&rsfc("xhgg")&">"

	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单位：</strong></td>"      
     response.write"<td width='88%' class='tdbg'><input type='text' disabled='disabled' value="&rsfc("dw")&"></td></tr>  "   
	 response.write"<input  name='kcgl_dw'  type='hidden' value="&rsfc("dw")&">"
   
   
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单价：</strong></td>"      
     response.write"<td width='88%' class='tdbg'><input type='text'disabled='disabled' value="&rsfc("dmoney")&">元</td></tr>  "   
 	 response.write"<input  name='kcgl_dmoney' type='hidden' value="&rsfc("dmoney")&">"
  
   	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数量：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_numb'  onBlur=""checkamoney()"" value="&rsfc("numb")&" ></td></tr>  "   
   	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>金额：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><div id=""checkamoney"" style=""display:none"" class=""ok""></div>元</td></tr>  "   
	 	
		 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>出库日期：</strong></td> "
   response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   response.write"<input name='kcgl_fcdate' type='text' value="&now()&" >"
   response.write"<a href='#' onClick=""popUpCalendar(this,kcgl_fcdate, 'yyyy-mm-dd'); return false;"">"
   response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf

	 
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='kcgl_bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='savefc'>  <input type='hidden' name='id' value='"&id&"'>   <input  type='submit' name='Submit' value=' 完 成 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
    rsfc.close
    set rsfc=nothing
end sub



sub savefc()    '保存出存信息
	  '保存到出库表中
  dim rsfc,sqlfc
  dim sscj
  dim rscheck,sqlcheck
  set rscheck=server.createobject("adodb.recordset")
  sqlcheck="select * from xc where id="&request("id")
  rscheck.open sqlcheck,connkc,1,3
  '如果现存表中的数量小于出库数量就返回，否则执行出库，如果现存表中数量等于出库数量，刚删除现存表中内容
  if rscheck("numb") < Cint(request("kcgl_numb")) then 
      response.write"<Script Language=Javascript>window.alert('出库数量大于现存数量');history.go(-1);</Script>"
  else
	  'd/保存信息到出库表中
	  set rsfc=server.createobject("adodb.recordset")
      sqlfc="select * from fc" 
      rsfc.open sqlfc,connkc,1,3
      rsfc.addnew
      'on error resume next
	  rsfc("wpid")=rscheck("wpid")
	  rsfc("class")=request("kcgl_class")
      rsfc("sscj")=request("kcgl_sscj")'源所属车间
      rsfc("qxtxt")=request("kcgl_qxtxt")'出库到的单位
	  rsfc("name")=Trim(request("kcgl_name"))
      rsfc("xhgg")=request("kcgl_xhgg")
      rsfc("dw")=request("kcgl_dw")
      rsfc("dmoney")=request("kcgl_dmoney")
      rsfc("numb")=request("kcgl_numb")'出库数量
      rsfc("amoney")=request("kcgl_dmoney")*request("kcgl_numb")'出库金额
      dim fcdate
	  fcdate=request("kcgl_fcdate")
	  if fcdate="" then fcdate=year(now())&"-"&month(now())&"-"&day(now())
	  rsfc("fc_year")=year(fcdate)
  	  rsfc("fc_month")=month(fcdate)
  	  rsfc("fc_day")=day(srdate)
	  rsfc("bz")=request("kcgl_bz")
      rsfc.update
      rsfc.close
      set rsfc=nothing


'增加入库表中.所属车间为去向
      dim sqlsr,rssr
	  set rssr=server.createobject("adodb.recordset")
      sqlsr="select * from sr" 
      rssr.open sqlsr,connkc,1,3
      rssr.addnew
      'on error resume next
	  rssr("wpid")=rscheck("wpid")
	  rssr("class")=request("kcgl_class")
      rssr("sscj")=request("kcgl_qxtxt")'出库到的车间
      rssr("lytxt")=sscjh(request("kcgl_sscj"))
	  rssr("name")=Trim(request("kcgl_name"))
      rssr("xhgg")=request("kcgl_xhgg")
      rssr("dw")=request("kcgl_dw")
      rssr("dmoney")=request("kcgl_dmoney")
      rssr("numb")=request("kcgl_numb")'出库的数量
      rssr("amoney")=request("kcgl_dmoney")*request("kcgl_numb")'出库的金额
	  dim srdate
	  srdate=request("kcgl_fcdate")
	 if srdate="" then srdate=year(now())&"-"&month(now())&"-"&day(now())
	  rssr("sr_year")=year(srdate)
  	  rssr("sr_month")=month(srdate)
  	  rssr("sr_day")=day(srdate)
	  rssr("bz")=request("kcgl_bz")
      rssr.update
      rssr.close
      set rssr=nothing
     end if 


     'c/出库去向车间增加或编辑显存表中.如果XC表中有同样YSID的品种,则增加原数量,如没有则新建,所属车间为表单中去向
     if request("kcgl_qxtxt")<>1000 then
	  dim rseditcj,sqleditcj 
	  set rseditcj=server.createobject("adodb.recordset")
      sqleditcj="select * from xc where wpid="&rscheck("wpid")&" and sscj="&request("kcgl_qxtxt")
      rseditcj.open sqleditcj,connkc,1,3
      if rseditcj.eof and rseditcj.bof then 
	      set rsfc=server.createobject("adodb.recordset")
          sqlfc="select * from xc" 
          rsfc.open sqlfc,connkc,1,3
          rsfc.addnew
          'on error resume next
	      rsfc("wpid")=rscheck("wpid")
		  rsfc("class")=request("kcgl_class")
          rsfc("sscj")=request("kcgl_qxtxt")'出库到的车间
          'rsfc("qxtxt")=request("kcgl_qxtxt")
	      rsfc("name")=Trim(request("kcgl_name"))
          rsfc("xhgg")=request("kcgl_xhgg")
          rsfc("dw")=request("kcgl_dw")
          rsfc("dmoney")=request("kcgl_dmoney")
          rsfc("numb")=request("kcgl_numb")'出库的数量
          rsfc("amoney")=request("kcgl_dmoney")*request("kcgl_numb")'出库的金额
	      rsfc("rcdate")=request("kcgl_fcdate")
		  rsfc("bz")=request("kcgl_bz")
          rsfc.update
          rsfc.close
          set rsfc=nothing
	  else
	    rseditcj("numb")=rseditcj("numb")+request("kcgl_numb")
	    rseditcj("amoney")=request("kcgl_dmoney")*rseditcj("numb")'
		rseditcj("rcdate")=request("kcgl_fcdate")
		rseditcj.update
      end if
	  rseditcj.close
      set rseditcj=nothing
	  
	  
	 
       'a\出库源车间如果现存表中数量大于出库数量，编辑现存表中源车间的数量
    if rscheck("numb")>Cint(request("kcgl_numb")) then 
	  dim rsedit,sqledit
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from xc where id="&request("id")
      rsedit.open sqledit,connkc,1,3
      rsedit("numb")=rsedit("numb")-request("kcgl_numb")
      rsedit("amoney")=request("kcgl_dmoney")*rsedit("numb")
      rsedit("rcdate")=request("kcgl_fcdate")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	else
	 if rscheck("numb")=Cint(request("kcgl_numb")) then 
	    dim rsdel,sqldel
	    set rsdel=server.createobject("adodb.recordset")
        sqldel="delete * from xc where id="&request("id")
       rsdel.open sqldel,connkc,1,3
     end if 
	end if   
	
	  'e\如果现存表中数量等于出库数量，刚删除现存表中内容
	  
  response.write"<Script Language=Javascript>history.go(-2)</Script>"
  rscheck.close
  set rscheck=nothing
  end if 
end sub


'用于库存分类名称显示
Function kcclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from kcclass where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connkc,1,1
    do while not rsname.eof
	    kcclass=rsname("name")
		rsname.movenext
	loop
	rsname.close
	set rsname=nothing
end Function 
Call CloseConn
%>