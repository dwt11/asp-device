<%@language=vbscript codepage=936 %>
<%
'Option Explicit
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
'on error resume next
url=geturl
dim keys,sscjid,title
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 

dim url,lb,brxx,sql,rs,record,pgsz,total,page,start,rowcount,ii
dim rsadd,sqladd,id,rsdel,sqldel,rsedit,sqledit
dim sqlscdate,rsscdate'上次周检时间
dim zjmonth '周检月份
Dwt.Out "<html>"& vbCrLf
Dwt.Out "<head>" & vbCrLf
Dwt.Out "<title>信息管理系统检验测量试验设备台账管理页</title>"& vbCrLf
Dwt.Out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.Out "<SCRIPT language=javascript>" & vbCrLf
Dwt.Out "function checkadd(){" & vbCrLf
Dwt.Out "if(document.form1.sscj.value==''){" & vbCrLf
Dwt.Out "alert('请选择所属单位！');" & vbCrLf
Dwt.Out "document.form1.sscj.focus();" & vbCrLf
Dwt.Out "return false;" & vbCrLf
Dwt.Out "}" & vbCrLf

'Dwt.Out "if(document.form1.zjtz_wh.value==''){" & vbCrLf
'Dwt.Out "alert('位号不能为空！');" & vbCrLf
'Dwt.Out "document.form1.zjtz_wh.focus();" & vbCrLf
'Dwt.Out "return false;" & vbCrLf
'Dwt.Out "}" & vbCrLf
Dwt.Out "}" & vbCrLf
Dwt.Out "</SCRIPT>" & vbCrLf
Dwt.Out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out"<script language=javascript src='/js/popselectdate.js'></script>"
Dwt.Out "</head>"& vbCrLf
Dwt.Out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
action=request("action")
select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "editd"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editd
  case "saveeditd"
    call saveeditd
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
  case "history"
    call history
  case "editinfo"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call editinfo
  case "saveeditinfo"
    call saveeditinfo
  case "delinfo"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call delinfo
	
end select	

Sub add()
	Dwt.Out"<br><br><br><form method='post' action='jycltz.asp' name='form1' onSubmit='javascript:return checkadd();'>"
	Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.Out"<tr class='title'><td height='22' colspan='2'>"
	Dwt.Out"<Div align='center'><strong>添加周检</Div></td>    </tr>"
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>所属单位： </strong></td>"      
	Dwt.Out"<td width='80%' class='tdbg'>"
	if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
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
    dwt.out "<select name='ssbz' size='1' >" & vbCrLf
    dwt.out "<option  selected>选择班组分类</option>" & vbCrLf
    dwt.out "</select></td></tr>  "  & vbCrLf
    dwt.out "<script><!--" & vbCrLf
    dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
    dwt.out "var group=new Array(groups)" & vbCrLf
    dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    dwt.out "group[i]=new Array()" & vbCrLf
    dwt.out "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   dwt.out "group["&rscj("levelid")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		else
		   dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		do while not rsbz.eof
		   dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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




    dwt.out "var temp=document.form1.ssbz" & vbCrLf
    dwt.out "function redirect(x){" & vbCrLf
    dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    dwt.out "temp.options[m]=null" & vbCrLf
    dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    dwt.out "}" & vbCrLf
    dwt.out "temp.options[0].selected=true" & vbCrLf
    dwt.out "}//--></script>" & vbCrLf



  else 	 
   dwt.out"<input name='sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   dwt.out"<input name='sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   dwt.out"<select name='ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  dwt.out"<option value='0'>车间</option>"
   else   
	  dwt.out"<option value='0'>车间</option>"
      do while not rs.eof
	     dwt.out"<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	  end if 
	 dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
	Dwt.Out"</td></tr>"& vbCrLf
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>类&nbsp;&nbsp;型：</strong></td> "
	Dwt.Out"<td><select name='zjtz_lx' size='1'>"
	dim sqlname,rsname
	sqlname="SELECT * from jycl_class "
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    do while not rsname.eof
		Dwt.Out "<option value='"&rsname("id")&"'>"&rsname("name")&"</option>"
		rsname.movenext
	loop
	end if 
	rsname.close
	set rsname=nothing
    Dwt.Out"</select></td></tr>"
	 
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>管理方式：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_glfs' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>出厂编号：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ccbh' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>生产产家：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_sccj' type='text'></td>    </tr>   "
	
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>规格型号：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ggxh' type='text'></td>    </tr>   "
	'Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>规格型号：</strong></td>"
	'Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ggxh' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量范围：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_clfw' type='text'></td>    </tr>   "
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>鉴定周期：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'><select name='zjtz_jdzq' size='1'>"
	Dwt.Out "<option value='12'>12个月</option>"
	Dwt.Out "<option value='24'>24个月</option>"
	Dwt.Out "<option value='36'>36个月</option>"
	Dwt.Out "<option value='0'>停用</option>"
	Dwt.Out "<option value='1'>不周检</option>"
	Dwt.Out "</select></td></tr>"
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>上次周检日期：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
    Dwt.out"<input name='zjtz_date'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'/>日常周检日期"
	
	
	Dwt.Out"<br/>请根据鉴定周期和下次鉴定时间来计算出一个模拟的上次周检日期</td>    </tr>   "

	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    Dwt.Out"<td width='80%' class='tdbg'><input type='text' name='zjtz_bz'></td></tr>  "   

	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='Submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.Out"</table></form>"
end Sub	

Sub saveadd()    
	  'on error resume next
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from jycltz" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("sscj")) 
      rsadd("ssbz")=Trim(Request("ssbz"))
      rsadd("ggxh")=request("zjtz_ggxh")
      rsadd("clfw")=request("zjtz_clfw")
      rsadd("jdzq")=cint(request("zjtz_jdzq"))
      rsadd("glfs")=request("zjtz_glfs")
      rsadd("ccbh")=request("zjtz_ccbh")
	  rsadd("sccj")=request("zjtz_sccj")
	  rsadd("class")=cint(request("zjtz_lx"))
          if request("zjtz_date")<>"" then
	  rsadd("sczjdate")=request("zjtz_date")
          end if
	  rsadd("bz")=request("zjtz_bz")
	  rsadd.update
	  rsadd.close
      set rsadd=nothing
	  
	  
	  
	  Dwt.Out"<Script Language=Javascript>location.href='jycltz.asp';</Script>"
end Sub

Sub saveeditd()    
      'on error resume next
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jycltz where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
      rsedit("sscj")=Trim(Request("sscj"))
	   if Request("ssbz")<>"" then 
      rsedit("ssbz")=Trim(Request("ssbz"))
	    end if
      rsedit("ggxh")=request("zjtz_ggxh")
	  rsedit("glfs")=request("zjtz_glfs")
      rsedit("ccbh")=request("zjtz_ccbh")
	  rsedit("sccj")=request("zjtz_sccj")
      rsedit("clfw")=request("zjtz_clfw")
      rsedit("jdzq")=cint(request("zjtz_jdzq"))
	  rsedit("class")=cint(request("zjtz_lx"))
      rsedit("bz")=request("zjtz_bz")
      if request("zjtz_date")<>"" then
      rsedit("sczjdate")=request("zjtz_date")
      end if
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  Dwt.Out"<Script Language=Javascript>history.go(-2)</Script>"
end Sub

Sub del()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from jycltz where id="&id
  rsdel.open sqldel,connzj,1,3
  set rsdel=nothing  
  
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete from jycl_info where zjtzid="&id
  rsdel.open sqldel,connzj,1,3
  
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end Sub
Sub delinfo()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from jycl_info where id="&id
  rsdel.open sqldel,connzj,1,3
  set rsdel=nothing  
  
  
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end Sub


Sub editd()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from jycltz where id="&id
   rsedit.open sqledit,connzj,1,1
   Dwt.Out"<br><br><br><form method='post' action='jycltz.asp' name='form1' onSubmit='javascript:return checkadd();'>"& vbCrLf
   Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   Dwt.Out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.Out"<Div align='center'><strong>编辑周检</Div></td>    </tr>"& vbCrLf
     Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>所属单位： </strong></td>"   & vbCrLf   
     Dwt.Out"<td width='80%' class='tdbg'>"& vbCrLf
     Dwt.Out"<input name='sscj' type='hidden' value="&rsedit("sscj")&">"& vbCrLf

	dim sqlbz,rsbz
	sqlbz="SELECT * from bzname where sscj="&rsedit("sscj")
   	set rsbz=server.createobject("adodb.recordset")
   	rsbz.open sqlbz,conn,1,1
   	Dwt.Out"<select name='ssbz' size='1'>"
   	if rsbz.eof and rsbz.bof then 
   		  Dwt.Out"<option value='0'>未添加班组</option>"& vbCrLf
   	else   
		  Dwt.Out"<option value='0'>车间</option>"
   	   do while not rsbz.eof
		     Dwt.Out"<option value='"&rsbz("id")&"'"
			 if rsedit("ssbz")=rsbz("id") then Dwt.Out " selected"
			 Dwt.Out">"&rsbz("bzname")&"</option>"& vbCrLf
		  rsbz.movenext
   	   loop
	end if 
	 Dwt.Out"</select>" & vbCrLf
 	 rsbz.close
 	 set rsbz=nothing
	 Dwt.Out"</td></tr>"& vbCrLf
	 Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>名&nbsp;&nbsp;称：</strong></td> "
	Dwt.Out"<td><select name='zjtz_lx' size='1'>"
	dim sqlname,rsname
	sqlname="SELECT * from jycl_class "
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    do while not rsname.eof
		Dwt.Out "<option value='"&rsname("id")&"'"
        if rsedit("class")=rsname("id") then Dwt.Out "selected"
		Dwt.Out ">"&rsname("name")&"</option>"
		rsname.movenext
	loop
	end if 
	rsname.close
	set rsname=nothing
    Dwt.Out"</select></td></tr>"
	 
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>管理方式：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_glfs' type='text' value="&rsedit("glfs")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>生产产家：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_sccj' type='text' value="&rsedit("sccj")&"></td>    </tr>   "
	 Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>出厂编号：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ccbh' type='text' value="&rsedit("ccbh")&"></td>    </tr>   "

    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>规格型号：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_ggxh' type='text' value="&rsedit("ggxh")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量范围：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjtz_clfw' type='text' value="&rsedit("clfw")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>鉴定周期：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><select name='zjtz_jdzq' size='1'>"
      Dwt.Out "<option value='12'"
      if rsedit("jdzq")=12 then Dwt.Out "selected"
	  Dwt.Out ">12个月</option>"
      Dwt.Out "<option value='24'"
      if rsedit("jdzq")=24 then Dwt.Out "selected"
	  Dwt.Out ">24个月</option>"
      Dwt.Out "<option value='36'"
      if rsedit("jdzq")=36 then Dwt.Out "selected"
	  Dwt.Out">36个月</option>"
      Dwt.Out "<option value='0'"
      if rsedit("jdzq")=0 then Dwt.Out "selected"
	  Dwt.Out">停用</option>"
      Dwt.Out "<option value='1'"
      if rsedit("jdzq")=1 then Dwt.Out "selected"
      Dwt.Out">不周检</option>"
       Dwt.Out "</select></td></tr>"
    
	    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>上次周检日期：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
     Dwt.out"<input name='zjtz_date' "
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("sczjdate")&"'/>日常周检日期"
	
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    Dwt.Out"<td width='80%' class='tdbg'><input type='text' name='zjtz_bz' value="&rsedit("bz")&"></td></tr>  "   
	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"& vbCrLf
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveeditd'> <input type='hidden' name='id' value='"&id&"'>      <input  type='Submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"& vbCrLf
	Dwt.Out"</table></form>"& vbCrLf
	       rsedit.close
       set rsedit=nothing
end Sub
Sub editinfo()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from jycl_info where id="&id
   rsedit.open sqledit,connzj,1,1
   Dwt.Out"<br><br><br><form method='post' action='jycltz.asp' name='form1' >"& vbCrLf
   Dwt.Out"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"& vbCrLf
   Dwt.Out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.Out"<Div align='center'><strong>编辑周检历史</strong></Div></td>    </tr>"& vbCrLf
   Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>周检日期：</strong></td>"
	Dwt.Out"<td width='80%' class='tdbg'>"
    Dwt.out"<br/><input name='zjtz_date' "
	dwt.out " onClick='new Calendar(0).show(this)' readOnly  value='"&request("zjdate")&"'/>日常周检日期"		
	Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>周检结果：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='zjinfo' type='text' value="&rsedit("zjinfo")&"></td>    </tr>   "
    Dwt.Out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备注：</strong></td>"
	 Dwt.Out"<td width='80%' class='tdbg'><input name='bz' type='text' value="&rsedit("bz")&"></td>    </tr>   "
	
	Dwt.Out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"& vbCrLf
	Dwt.Out"<input name='action' type='hidden' id='action' value='saveeditinfo'> <input type='hidden' name='id' value='"&id&"'>      <input  type='Submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"& vbCrLf
	Dwt.Out"</table></form>"& vbCrLf
	       rsedit.close
       set rsedit=nothing
end Sub
sub saveeditinfo()
	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jycl_info where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connzj,1,3
	     rsedit("zjdate")=request("zjtz_date")
	     zjyear=year(request("zjtz_date"))
		 zjmonth=month(request("zjtz_date"))
      zjtzid=rsedit("zjtzid")
	  rsedit("bz")=request("bz")
      rsedit("zjinfo")=request("zjinfo")
	  rsedit.update
      set rsedit=nothing
	  
	  	 	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from jycltz where id="&zjtzid
      rsedit.open sqledit,connzj,1,3
	     rsedit("sczjdate")=request("zjtz_date")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
  Dwt.Out"<Script Language=Javascript>history.go(-1)</Script>"
end sub
Sub history()

    sql="SELECT * from jycltz where id="&request("id")&" ORDER BY id DESC"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzj,1,1
    if rs.eof and rs.bof then 
        Dwt.Out "<p align='center'>未找到内容</p>" 
    else
		Dwt.Out "<Div style='left:6px;'>"& vbCrLf
		Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
		Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&rs("class")&"  周检历史</span>"& vbCrLf
		Dwt.Out "     </Div>"& vbCrLf		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>单位</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>名称</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>管理方式</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>生产产家</Div></td>" & vbCrLf

        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>出厂编号</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>型号</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>测量范围</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>鉴定周期</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))&"</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("glfs")&"&nbsp;</td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'>"&rs("sccj")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("ccbh")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
         Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("jdzq")&"&nbsp;</Div></td>" & vbCrLf
        Dwt.Out "</tr></table>" & vbCrLf
	  sscjid=rs("sscj")
	end if
	
	
    rs.close
    set rs=nothing
	
	sqlscdate="SELECT * from jycl_info where zjtzid="&request("id")&" ORDER BY id DESC"
    set rsscdate=server.createobject("adodb.recordset")
    rsscdate.open sqlscdate,connzj,1,1
    if rsscdate.eof and rsscdate.bof then 
        message("没有以前的周检记录")
    else
         record=rsscdate.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsscdate.PageSize = Cint(PgSz) 
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
           rsscdate.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsscdate.PageSize
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>序号</Div></td>" & vbCrLf
        Dwt.Out "      <td class='x-td'  ><Div class='x-grid-hd-text'>周检日期</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>鉴定结果</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>备注</Div></td>" & vbCrLf
        Dwt.Out "      <td  class='x-td' ><Div class='x-grid-hd-text'>选项</Div></td>" & vbCrLf
        Dwt.Out "    </tr>" & vbCrLf
		   do while not rsscdate.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
        Dwt.Out "      <td class='x-td' ><Div align=""center"">"&xh_id&"</Div></td>" & vbCrLf
                      zjdate=rsscdate("zjdate")
		Dwt.Out "      <td  class='x-td'>"&zjdate&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("zjinfo")&"&nbsp;</td>" & vbCrLf
        Dwt.Out "      <td  class='x-td'>"&rsscdate("bz")&"&nbsp;</td>" & vbCrLf
		if session("levelclass")=sscjid or session("levelclass")=0 then 
			Dwt.Out "<td  class='x-td'><a href=jycltz.asp?action=editinfo&id="&rsscdate("id")&">编辑</a>&nbsp;"
			Dwt.Out "<a href=jycltz.asp?action=delinfo&id="&rsscdate("id")&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a></td>"
		 else
			Dwt.Out "&nbsp;"
		 end if 
 
			 RowCount=RowCount-1
          rsscdate.movenext
          loop
        Dwt.Out "</table>" & vbCrLf
       url="jycltz.asp?action=history&id="&request("id")
	   call showpage(page,url,total,record,PgSz)
	   Dwt.Out "</Div>"
	   end if
	   Dwt.Out "</Div>"
	          rsscdate.close
	         Dwt.Out "<br><table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""><tr><td>" 
			Dwt.Out "<input name='Cancel' type='button' id='Cancel' value=' 返  回 ' onClick="";history.back()"" style='cursor:hand;'></td></tr></table>"

end Sub







Sub main()
	'dim sql,rsjxjl,title
	sql="SELECT * from jycltz"
	if keys<>"" then 
		sql=sql&" where class like '%" &keys& "%' "
		title="-搜索 "&keys
	end if 
	if sscjid<>"" then
		sql=sql&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
	end if 
	sql=sql&" ORDER BY sscj aSC "
	
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>检验测量试验设备周检台账"&title&"</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf


call search()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
		Dwt.Out "<p align='center'>未添加内容</p>" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>序号</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>名称</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>类型</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>管理方式</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>生产产家</Div></td>" & vbCrLf

		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>出厂编号</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>型号</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>测量范围</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>鉴定周期</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>上次鉴定</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>下次鉴定</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>备注</Div></td>" & vbCrLf
		'Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>选项</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		record=rs.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
				   PgSz=Trim(Request("PgSz"))
			   end if 
			   rs.PageSize = Cint(PgSz) 
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
			   rs.absolutePage = page
			   start=PgSz*Page-PgSz+1
			   rowCount = rs.PageSize
		do while not rs.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.Out "     <td  CLASS='X-TD'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
					Dwt.Out "      <td class='x-td' ><Div align=""center"">"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))

call edit(rs("id"),rs("sscj"))
DWT.OUT "</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("glfs")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("sccj")&"&nbsp;</td>" & vbCrLf					
					Dwt.Out "      <td  class='x-td'>"&rs("ccbh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("clfw")&"&nbsp;</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=""center"">"&rs("jdzq")&"&nbsp;</Div></td>" & vbCrLf
	
					dim jdzq  '检定周期判断
					dim jdinfo
					dim jdyear '检定周期换算为年
					jdzq=rs("jdzq")/12
					
			'上次周检日期
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"				   
			     Dwt.out rs("sczjdate")
			     Dwt.out "</Div></td>"
			'下次周检日期
			     Dwt.Out "<td  class='x-td'><Div align=""center"">"
			     Dwt.out year(rs("sczjdate"))+jdzq&"-"&month(rs("sczjdate"))
			     Dwt.out "</Div></td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'>"&rs("bz")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
					Dwt.Out "</Div></td></tr>" & vbCrLf
					 RowCount=RowCount-1
			  rs.movenext
			  loop
			Dwt.Out "</table>" & vbCrLf
		   if sscjid<>"" or keys<>"" then 
		       call showpage(page,url,total,record,PgSz)
			else
		       call showpage1(page,url,total,record,PgSz)
           end if 
		   Dwt.Out "</Div>"
		   end if
		   Dwt.Out "</Div>"		   
		   rs.close
		   set rs=nothing
end Sub
Dwt.Out "</body></html>"

'用于分类名称显示
Function zjclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from jycl_class where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	else
	    zjclass=rsname("name")
	end if 
	rsname.close
	set rsname=nothing
end Function

Sub edit(id,sscj)
    Dwt.Out " <a href=jycltz.asp?action=history&id="&id&">史</a>&nbsp;"
if session("levelclass")=sscj or session("levelclass")=0 then 
    Dwt.Out "<a href=jycltz.asp?action=editd&id="&id&">编</a>&nbsp;"
	Dwt.Out "<a href=jycltz.asp?action=del&id="&id&" onClick=""return confirm('此操作会删除该表所有的周检记录，确定要删除此记录吗？');"">删</a>"
 else
    Dwt.Out "&nbsp;"
 end if 
end Sub




Sub search()
	dim sqlcj,rscj
    Dwt.Out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.Out "<form method='Get' name='SearchForm' action='jycltz.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.Out "<a href=""jycltz.asp?action=add"">添加周检</a>"
	Dwt.Out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if keys<>"" then 
	 Dwt.Out "value='"&keys&"'"
    	Dwt.Out ">" & vbCrLf
    else
	 Dwt.Out "value='输入搜索的位号'"
	 	Dwt.Out " onblur=""if(this.value==''){this.value='输入搜索的位号'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	Dwt.Out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	
	Dwt.Out "<select id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.Out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			Dwt.Out"<option value='jycltz.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then Dwt.Out" selected"

			Dwt.Out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		Dwt.Out "     </select>	" & vbCrLf
	
	
    Dwt.Out "</form></Div></Div>" & vbCrLf
end Sub





Call Closeconn
%>