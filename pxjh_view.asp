<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
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
'20080303数据库增加字段,sjjhtbdate\sjzjtbdate保存实际添写的时间,只有王茜可看
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>信息管理系统培训管理内容显示</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form1.pxjh_sscj.value==''){" & vbCrLf
Dwt.out "      alert('请选择所属车间！');" & vbCrLf
Dwt.out "   document.form1.pxjh_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "        //判断日期(长日期格式为:2006-04-03)的正则" & vbCrLf
Dwt.out "var dval1=document.form1.pxjh_date.value;" & vbCrLf
Dwt.out "var r=/^(\d{0,4})-(0{0,1}[1-9]|1[0-2])-(0{0,1}[1-9]|[1-2]\d|3[0-1])$/;" & vbCrLf
Dwt.out "if(!r.test(dval1)){" & vbCrLf
Dwt.out "    alert('输入日期错误');" & vbCrLf
Dwt.out "    document.form1.pxjh_date.focus();" & vbCrLf
Dwt.out "    return false;}" & vbCrLf
Dwt.out "else{" & vbCrLf
Dwt.out "    var r1=/^0{0,4}$/;" & vbCrLf
Dwt.out "    if(r1.test(RegExp.$1)){ " & vbCrLf
Dwt.out "           alert('年份不能为0');" & vbCrLf
Dwt.out "             document.form1.pxjh_date.focus();" & vbCrLf
Dwt.out "              return false;}" & vbCrLf
Dwt.out "              var r2=/1[02]|0{0,1}[13578]/;" & vbCrLf' //小月
Dwt.out "              if(!r2.test(RegExp.$2)){" & vbCrLf
Dwt.out "                     if(parseInt(RegExp.$2)==2){" & vbCrLf
Dwt.out "                            if(parseInt(RegExp.$1)%4==0){" & vbCrLf
Dwt.out "                                 if(parseInt(RegExp.$3)>29){" & vbCrLf
Dwt.out "                                       alert('闰年2月只有29天');" & vbCrLf
Dwt.out "                                       document.form1.pxjh_date.focus();" & vbCrLf
Dwt.out "                                       return false;}}" & vbCrLf
Dwt.out "                             else{" & vbCrLf
Dwt.out "                                 if(parseInt(RegExp.$3)>28){" & vbCrLf
Dwt.out "                                       alert('2月只有28天');" & vbCrLf
Dwt.out "                                       document.form1.pxjh_date.focus();" & vbCrLf
Dwt.out "                                       return false;}}}" & vbCrLf
Dwt.out "                             else{" & vbCrLf
Dwt.out "                                if(parseInt(RegExp.$3)>30){" & vbCrLf
Dwt.out "                                        alert('小月只有30天');" & vbCrLf
Dwt.out "                                        document.form1.pxjh_date.focus();" & vbCrLf
Dwt.out "                                        return false;}}}}" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

if request("action")="pxjh" then call pxjh()
if request("action")="wc" then call wc()
if request("action")="savewc" then call savewc()
if request("action")="addpxjh" then call addpxjh()
if request("action")="saveaddpxjh" then call saveaddpxjh()
if request("action")="del" then call del()
if request("action")="edit" then call edit()
if request("action")="saveedit" then saveedit()
sub addpxjh()
dim ii
dim rscj,sqlcj,rsbz,sqlbz,sql,rs
   Dwt.out"<br><br><br><form method='get' action='pxjh_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='20' colspan='2'>"
   Dwt.out"<Div align='center'><strong>添加培训报表-计划</strong></Div></td>    </tr>"
	Dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属车间： </strong></td>"& vbCrLf      
    Dwt.out"<td width='80%' class='tdbg'>"& vbCrLf
  if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	Dwt.out"<select name='pxjh_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
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
    Dwt.out "<select name='pxjh_ssbz' size='1' >" & vbCrLf
    Dwt.out "<option  selected>选择班组分类</option>" & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form1.pxjh_sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		else
		   Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		do while not rsbz.eof
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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




    Dwt.out "var temp=document.form1.pxjh_ssbz" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//--></script>" & vbCrLf



  else 	 
   Dwt.out"<input name='pxjh_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   Dwt.out"<input name='pxjh_sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   Dwt.out"<select name='pxjh_ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  Dwt.out"<option value='0'>车间</option>"
   else   
	  Dwt.out"<option value='0'>车间</option>"
      do while not rs.eof
	     Dwt.out"<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	  end if 
	 Dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
    Dwt.out"</td></tr>  "  	 

   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训日期：</strong></td> "
   Dwt.out"<td width='80%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   Dwt.out"<input name='pxjh_date' type='text' value="&now()&" >"
   Dwt.out"<a href='#' onClick=""popUpCalendar(this, pxjh_date, 'yyyy-mm-dd'); return false;"">"
   Dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 Dwt.out"<strong>培训内容摘要：</strong></td>"
	 Dwt.out"<td width='80%' class='tdbg'><input name='pxjh_body' type='text'  size=""50""></td>    </tr>   "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训对象：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_pxdx'></td></tr> "
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>计划人次：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_numb'  onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;""></td></tr> "
	 
	' Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训形式：</strong></td> "
'	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_xs'></td></tr> "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训课时：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'>"
	Dwt.out"<select name='pxjh_ks' size='1'>"
	Dwt.out"<option value='1'>1h</option>"
    Dwt.out"<option value='2'>2h</option>"
    Dwt.out"<option value='3'>3h</option>"
    Dwt.out"<option value='4'>4h</option>"
    Dwt.out"<option value='5'>5h</option>"
    Dwt.out"</select></td></tr>  "  	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>授课人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_skrname'></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备注：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_bz'></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>添报人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_tbrname'></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>单位主管：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_zgname'></td></tr> "
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveaddpxjh'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
end sub	

sub saveaddpxjh()    
	  dim year1,month1,day1'保存\
	  dim rsadd,sqladd
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from pxjh" 
      rsadd.open sqladd,conne,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("pxjh_sscj"))
      rsadd("ssbz")=Trim(Request("pxjh_ssbz"))
      year1=year(Trim(Request("pxjh_date")))
	  month1=month(Trim(Request("pxjh_date")))
	  day1=day(Trim(Request("pxjh_date")))
	  if len(month1)<>2 then month1="0"&month1
	  rsadd("day")=day1
      rsadd("month")=month1
	  rsadd("year")=year1
	  rsadd("tbrname")=request("pxjh_tbrname")
	  rsadd("zgname")=request("pxjh_zgname")
      rsadd("body")=Trim(request("pxjh_body"))
      rsadd("numb")=request("pxjh_numb")
      rsadd("pxdx")=request("pxjh_pxdx")
      rsadd("ks")=request("pxjh_ks")
      rsadd("skrname")=request("pxjh_skrname")
      rsadd("tbdate")=year(now())&"-"&month(now())&"-"&day(now())
      rsadd("sjjhtbdate")=now()
      rsadd("bz")=request("pxjh_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub



sub saveedit()    
	  dim year1,month1,day1'保存\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from pxjh where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conne,1,3
      rsedit("sscj")=Trim(Request("pxjh_sscj"))
      rsedit("ssbz")=Trim(Request("pxjh_ssbz"))
      year1=year(Trim(Request("pxjh_date")))
	  month1=month(Trim(Request("pxjh_date")))
	  day1=day(Trim(Request("pxjh_date")))
	  if len(month1)<>2 then month1="0"&month1
	  rsedit("day")=day1
      rsedit("month")=month1
	  rsedit("year")=year1
	  rsedit("tbrname")=request("pxjh_tbrname")
	  rsedit("zgname")=request("pxjh_zgname")
      rsedit("body")=ReplaceBadChar(request("pxjh_body"))
      rsedit("numb")=request("pxjh_numb")
      rsedit("pxdx")=request("pxjh_pxdx")
      rsedit("ks")=request("pxjh_ks")
      rsedit("skrname")=request("pxjh_skrname")
      rsedit("tbdate")=year(now())&"-"&month(now())&"-"&day(now())
      rsedit("bz")=request("pxjh_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub



sub edit()

   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from pxjh where id="&id
   rsedit.open sqledit,conne,1,1

   Dwt.out"<br><br><br><form method='get' action='pxjh_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>编辑培训计划</strong></Div></td>    </tr>"
	
	Dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属车间： </strong></td>"& vbCrLf      
    Dwt.out"<td width='80%' class='tdbg'>"& vbCrLf
    Dwt.out"<input name=""pxjh_sscj"" value="&sscjh(rsedit("sscj"))&" type='text' disabled='disabled' >"& vbCrLf
     Dwt.out"<input name='pxjh_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	
	Dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属班组： </strong></td>"& vbCrLf      
    Dwt.out"<td width='80%' class='tdbg'>"& vbCrLf
	if rsedit("ssbz")=0 then
  	   ssbz="车间"
	else
	   ssbz=ssbzh(rsedit("ssbz"))
	end if    
    Dwt.out"<input name=""pxjh_ssbz"" value="&ssbz&" type='text' disabled='disabled' >"& vbCrLf
     Dwt.out"<input name='pxjh_ssbz' type='hidden' value="&rsedit("ssbz")&"></td></tr>"& vbCrLf

   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训日期：</strong></td> "
   Dwt.out"<td width='80%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   Dwt.out"<input name='pxjh_date' type='text' value="&rsedit("year")&"-"&rsedit("month")&"-"&rsedit("day")&" >"
   Dwt.out"<a href='#' onClick=""popUpCalendar(this, pxjh_date, 'yyyy-mm-dd'); return false;"">"
   Dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 Dwt.out"<strong>培训内容摘要：</strong></td>"
	 Dwt.out"<td width='80%' class='tdbg'><input name='pxjh_body' type='text'  size=""50"" value='"&rsedit("body")&"'></td>    </tr>   "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训对象：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_pxdx'  value='"&rsedit("pxdx")&"'></td></tr> "
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>计划人次：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_numb'  value='"&rsedit("numb")&"' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;""></td></tr> "
	 
	' Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训形式：</strong></td> "
'	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_xs'></td></tr> "
	 
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训课时：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'>"
	Dwt.out"<select name='pxjh_ks' size='1'>"
	
	Dwt.out"<option value='1'"
	if rsedit("ks")=1 then Dwt.out" selected"
	Dwt.out">1h</option>"
	
    Dwt.out"<option value='2'"
	if rsedit("ks")=2 then Dwt.out" selected"
	Dwt.out">2h</option>"
	
    Dwt.out"<option value='3'"
	if rsedit("ks")=3 then Dwt.out" selected"
	Dwt.out">3h</option>"
	
    Dwt.out"<option value='4'"
	if rsedit("ks")=4 then Dwt.out" selected"
	Dwt.out">4h</option>"
	
    Dwt.out"<option value='5'"
	if cint(rsedit("ks"))=5 then Dwt.out" selected"
	Dwt.out">5h</option>"
    
	Dwt.out"</select></td></tr>  "  	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>授课人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_skrname' value="&rsedit("skrname")&"></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备注：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_bz' value="&rsedit("bz")&"></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>添报人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_tbrname' value="&rsedit("tbrname")&"></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>单位主管：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_zgname' value="&rsedit("zgname")&"></td></tr> "
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub


sub savewc()    
	  dim year1,month1,day1'保存\
	  dim rsedit,sqledit
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from pxjh where id="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,conne,1,3
      rsedit("sjnumb")=request("pxjh_sjnumb")
      rsedit("sjks")=request("pxjh_sjks")
      rsedit("hgl")=request("pxjh_hgl")
      rsedit("pxl")=request("pxjh_pxl")
      rsedit("bz")=request("pxjh_bz")
      rsedit("tbrname")=request("pxjh_tbrname")
      rsedit("sjzjtbdate")=now()
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub wc()

   dim id,rsedit,sqledit,ssbz
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from pxjh where id="&id
   rsedit.open sqledit,conne,1,1

   Dwt.out"<br><br><br><form method='get' action='pxjh_view.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   Dwt.out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"
   Dwt.out"<Div align='center'><strong>培训报表-月底总结</strong></Div></td>    </tr>"
	
	Dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属车间： </strong></td>"& vbCrLf      
    Dwt.out"<td width='80%' class='tdbg'>"& vbCrLf
    Dwt.out"<input name=""pxjh_sscj"" value="&sscjh(rsedit("sscj"))&" type='text' disabled='disabled' >"& vbCrLf
     'Dwt.out"<input name='pxjh_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	
	Dwt.out"<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属班组： </strong></td>"& vbCrLf      
    Dwt.out"<td width='80%' class='tdbg'>"& vbCrLf
	if rsedit("ssbz")=0 then
  	   ssbz="车间"
	else
	   ssbz=ssbzh(rsedit("ssbz"))
	end if    
    Dwt.out"<input name=""pxjh_ssbz"" value="&ssbz&" type='text' disabled='disabled' >"& vbCrLf
     'Dwt.out"<input name='pxjh_ssbz' type='hidden' value="&rsedit("ssbz")&"></td></tr>"& vbCrLf

   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训计划日期：</strong></td> "
   Dwt.out"<td width='80%' class='tdbg'>"
   Dwt.out"<input name='pxjh_date' type='text' value="&rsedit("year")&"-"&rsedit("month")&"-"&rsedit("day")&"  disabled='disabled'>"
   'Dwt.out"<a href='#' onClick=""popUpCalendar(this, pxjh_date, 'yyyy-mm-dd'); return false;"">"
   'Dwt.out"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 Dwt.out"<strong>培训内容摘要：</strong></td>"
	 Dwt.out"<td width='80%' class='tdbg'><input name='pxjh_body' type='text'  size=""50""  disabled='disabled' value='"&rsedit("body")&"'></td>    </tr>   "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训对象：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_numb'  disabled='disabled' value="&rsedit("pxdx")&"></td></tr> "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>计划人数：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_xs'  disabled='disabled' value="&rsedit("numb")&"></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>实际人数：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_sjnumb'  onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" value="&rsedit("sjnumb")&"></td></tr> "
	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>计划课时：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_ks'  disabled='disabled' value="&rsedit("ks")&"   disabled='disabled' >"
	Dwt.out"</td></tr>  "  	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>实际课时：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'>"
		Dwt.out"<select name='pxjh_sjks' size='1'>"
	Dwt.out"<option value='1'>1h</option>"
    Dwt.out"<option value='2'>2h</option>"
    Dwt.out"<option value='3'>3h</option>"
    Dwt.out"<option value='4'>4h</option>"
    Dwt.out"<option value='5'>5h</option>"
    Dwt.out"</select></td></tr>  "  	 
   Dwt.out"</td></tr>  "  	 
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>培训率：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_pxl' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" value="&rsedit("pxl")&"   >"
	Dwt.out"</td></tr>  "  	 
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>合格率：</strong></td> "
    Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_hgl' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" value="&rsedit("hgl")&"  >"
	Dwt.out"</td></tr>  "  	 
	
	
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>授课人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_skrname' value="&rsedit("skrname")&"  disabled='disabled'  ></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备注：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_bz' value="&rsedit("bz")&"></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>添报人：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_tbrname' value="&rsedit("tbrname")&"   ></td></tr> "
	 Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>单位主管：</strong></td> "
	 Dwt.out"<td width='80%' class='tdbg'><input type='text' name='pxjh_zgname' value="&rsedit("zgname")&"   disabled='disabled' ></td></tr> "
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='savewc'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub


sub del()
 dim rsdel,sqldel
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from pxjh where id="&request("id")
  rsdel.open sqldel,conne,1,3
  Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
set rsdel=nothing  

end sub


sub pxjh()
Dwt.out "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
Dwt.out " <tr class='topbg'>"& vbCrLf

Dwt.out "   <td height='22' colspan='2' align='center'><strong><Div align=center><strong>"&sscjh(request("sscj"))&request("year")&"年"&request("month")&"月份培训计划</strong></Div></strong></td>"& vbCrLf
Dwt.out "  </tr>  "& vbCrLf
Dwt.out "</table>"& vbCrLf

   dim rspxjh,sqlpxjh,rs,sql
   '显示车间级的培训计划
      sqlpxjh="SELECT * from pxjh where ssbz=0 and sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rspxjh=server.createobject("adodb.recordset")
      rspxjh.open sqlpxjh,conne,1,1
      if rspxjh.eof and rspxjh.bof then 
             Dwt.out "<p align='center'>未添加车间培训计划</p>" 
          else
             Dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
             Dwt.out " <tr class=""title""><td colspan=14 >&nbsp;&nbsp;&nbsp;  单位："&sscjh(request("sscj"))&"&nbsp;"&ssbzh(rspxjh("ssbz"))
             Dwt.out "</td></tr>"
             Dwt.out "<tr class=""title""><td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>时间</Div></td>"
             Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训内容摘要</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训对象</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划人数</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>实际人数</Div></td>"
            ' Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训形式</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划课时</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>实际课时</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训率</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>合格率</Div></td>"
             Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>授课人</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>备注</Div></td>"
			               '如果登录用户是王则显示实际日期

			if session("userid")=80 or session("userid")=108 then
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>选项</Div></td>"
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划添报日期</Div></td>"
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>总结添报日期</Div></td></tr>"
            else
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>选项</Div></td></tr>"
 
            end if
              do while not rspxjh.eof
                 Dwt.out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
				 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("month")&"."&rspxjh("day")&"</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px"">"&rspxjh("body")&"&nbsp;</td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("pxdx")&"&nbsp;</Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("numb")&"&nbsp;</Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjnumb")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("ks")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjks")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("pxl")&"%</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("hgl")&"%</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("skrname")&"&nbsp; </Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px"">车间"&rspxjh("bz")&"&nbsp;</td>"
                 Dwt.out "      <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"
                 call editdel(rspxjh("id"),rspxjh("sscj"),"pxjh_view.asp?action=edit&id=","pxjh_view.asp?action=del&id=")
               '如果登录用户是王则显示实际日期
               if session("userid")=80 or session("userid")=108 then
				 Dwt.out "<a href=pxjh_view.asp?action=wc&id="&rspxjh("id")&">完成</a></Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjjhtbdate")&"&nbsp; </Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjzjtbdate")&"&nbsp; </Div></td></tr>"
               else  
				 Dwt.out "<a href=pxjh_view.asp?action=wc&id="&rspxjh("id")&">完成</a></Div></td></tr>"
               end if  
                 rspxjh.movenext
		      loop
          end if 
		  Dwt.out "  </tr></table><br>"
		  
'显示各车间所属班组培训		  
 sql="SELECT * from bzname where sscj="&request("sscj")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   do while not rs.eof
      sqlpxjh="SELECT * from pxjh where ssbz="&rs("id")&" and month="&request("month")&" and year="&request("year")
      set rspxjh=server.createobject("adodb.recordset")
      rspxjh.open sqlpxjh,conne,1,1
      if rspxjh.eof and rspxjh.bof then 
             Dwt.out "<p align='center'>未添加"&ssbzh(rs("id"))&"培训计划</p>" 
          else
             Dwt.out "<br><table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
             Dwt.out "<tr class=""title""><td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>时间</Div></td>"
             Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训内容摘要</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训对象</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划人数</Div></td>"
             Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>实际人数</Div></td>"
            ' Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训形式</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划课时</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>实际课时</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>培训率</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>合格率</Div></td>"
             Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>授课人</Div></td>"
             Dwt.out " <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>备注</Div></td>"
			               '如果登录用户是王则显示实际日期

			if session("userid")=80 or session("userid")=108 then
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>选项</Div></td>"
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>计划添报日期</Div></td>"
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>总结添报日期</Div></td></tr>"
            else
			 Dwt.out "<td  style=""border-bottom-style: solid;border-width:1px""><Div align=center>选项</Div></td></tr>"
 
            end if
              do while not rspxjh.eof
                 Dwt.out "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
				 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("month")&"."&rspxjh("day")&"</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px"">"&rspxjh("body")&"&nbsp;</td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("pxdx")&"&nbsp;</Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("numb")&"&nbsp;</Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjnumb")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("ks")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjks")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("pxl")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("hgl")&"&nbsp;</Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("skrname")&"&nbsp; </Div></td>"
                 Dwt.out "<td   style=""border-bottom-style: solid;border-width:1px"">"&ssbzh(rs("id"))&rspxjh("bz")&"&nbsp;</td>"
                 Dwt.out "      <td   style=""border-bottom-style: solid;border-width:1px""><Div align=center>"
                 call editdel(rspxjh("id"),rspxjh("sscj"),"pxjh_view.asp?action=edit&id=","pxjh_view.asp?action=del&id=")
               '如果登录用户是王则显示实际日期
               if session("userid")=80 or session("userid")=108 then
				 Dwt.out "<a href=pxjh_view.asp?action=wc&id="&rspxjh("id")&">完成</a></Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjjhtbdate")&"&nbsp; </Div></td>"
                 Dwt.out "<td    style=""border-bottom-style: solid;border-width:1px""><Div align=center>"&rspxjh("sjzjtbdate")&"&nbsp; </Div></td></tr>"
               else  
				 Dwt.out "<a href=pxjh_view.asp?action=wc&id="&rspxjh("id")&">完成</a></Div></td></tr>"
               end if  
                 rspxjh.movenext
		      loop
          end if 
		  Dwt.out "  </tr></table><br>"
        rs.movenext
  loop
  rs.close
  set rs=nothing
  rspxjh.close
  set rspxjh=nothing
end sub


Dwt.out "</body></html>"
'************************************
'各车间登录后显示对应的编辑和删除
'*************************************

Call CloseConn
%>