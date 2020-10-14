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

dim url,lb,brxx,sqlqptjtz,rsqptjtz,record,pgsz,total,page,start,rowcount,ii
dim rsadd,sqladd,id,rsdel,sqldel,rsedit,sqledit,qptyqk
'url=geturl
dim keys,sscjid,ssghid,onnumb,sqld,rsd,sqlcj,rscj

keys=trim(request("keyword")) 
sscjid=trim(request("sscj"))
qptyqk=trim(request("styqk"))
dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统气瓶统计台帐管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out "if(document.form1.qptjtz_sscj.value==''){" & vbCrLf
dwt.out "alert('请选择使用单位！');" & vbCrLf
dwt.out "document.form1.qptjtz_sscj.focus();" & vbCrLf
dwt.out "return false;" & vbCrLf
dwt.out "}" & vbCrLf

dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "<SCRIPT language=javascript>" & vbCrLf
dwt.out "function checkadd(){" & vbCrLf
dwt.out "if(document.form1.qptjtz_sscj.value==''){" & vbCrLf
dwt.out "alert('请选择使用单位！');" & vbCrLf
dwt.out "document.form1.qptjtz_sscj.focus();" & vbCrLf
dwt.out "return false;" & vbCrLf
dwt.out "}" & vbCrLf

dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf

dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0' onload=""Javascript:document.form1.input1.focus();"">"& vbCrLf

action=request("action")

select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add
  case "saveadd"
    call saveadd
  case "edit"
	if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"
    call saveedit
  case "del"
    if truepagelevelh(session("groupid"),3,session("pagelevelid")) then call del
  case ""
	if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	

sub add()
  Dwt.out"<script type=""text/javascript"" src=""js/checkbh.js""></script>"&vbcrlf
 '新增用户

dim rscj,sqlcj
   dwt.out"<br><br><br><form method='post' action='qptjtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table  border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>添加气瓶统计台账</strong></div></td>    </tr>"
   dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>使用单位： </strong></td>"      
   dwt.out"<td  class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='qptjtz_sscj' size='1'>"
    dwt.out"<option >请选择使用单位</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
    dwt.out"<input name='qptjtz_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    dwt.out"<input name='qptjtz_sscj' type='hidden' value="&session("levelclass")&"></td></tr>"& vbCrLf
 end if 
 	 Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 Dwt.out"<strong>编&nbsp;&nbsp;号：</strong></td>"
	 Dwt.out"<td width='88%' class='tdbg'>"
	 Dwt.out "<input name='qptjtz_bh' type='text' id='input1' onblur='return myuser()' />"
	 Dwt.out "<span id='sps1'></span> "
	 Dwt.out "</td>    </tr>   "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>标气名称：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_bqname' ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>体&nbsp;&nbsp;积：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qptj' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>压&nbsp;&nbsp;力：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyl' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>成份含量：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpcf' ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>样品编号：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_ypbh' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>余&nbsp;&nbsp;气：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyq' ></td></tr> "
	 	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>用&nbsp;&nbsp;途：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_yt' ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>状态：</strong></td>"
	dwt.out"<td><select name='qptjtz_tyqk' size='1'>"
	dwt.out"<option value='1'>在用</option>"
	dwt.out"<option value='2'>待换</option>"
	dwt.out"<option value='3'>退库</option>"
    dwt.out"</select></td></tr>"


	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>定值日期</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='qptjtz_scdata' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
	dwt.out"</td></tr>  "   
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>有效期：</strong></td> "
	Dwt.Out"<td width='80%' class='tdbg'>"	

    dwt.out outdatadict ("qptjtz_yxq","有效期",onnumb)	 
    dwt.out "</td></tr>"
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>领用日期</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='qptjtz_lydata' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value="&date()&" >"
    dwt.out "</td></tr>"
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>存放地点：</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_cfdd'></td></tr>  "  
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>供气厂家：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_sccj'></td></tr> "
	 

   	 dwt.out"<tr><td width='12%' align='right' class='tdbg'><strong>到期日期</strong></td>"      
	 Dwt.out "<td width='88%' class='tdbg'><input name='qptjtz_dqdata' type='text' id='input6' onFocus='return addrdata()'/>&nbsp;<span >点击自动更新</span>"

	dwt.out"</td></tr>  "   	
  

	dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out"<td  class='tdbg'><input type='text' name='qptjtz_bz'></td></tr>  "   

	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
end sub	

sub saveadd()    
	  on error resume next
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from qptjtz" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
	  rsadd("sscj")=Trim(Request("qptjtz_sscj"))
      rsadd("bh")=request("qptjtz_bh")
      rsadd("bqname")=request("qptjtz_bqname") 
      rsadd("qptj")=Trim(request("qptjtz_qptj"))
      rsadd("qpyl")=request("qptjtz_qpyl")
      rsadd("qpcf")=request("qptjtz_qpcf")
	  rsadd("ypbh")=request("qptjtz_ypbh")
	  rsadd("qpyq")=request("qptjtz_qpyq")
	  rsadd("yt")=request("qptjtz_yt")
      rsadd("scdata")=request("qptjtz_scdata")
      rsadd("yxq")=request("qptjtz_yxq")
      rsadd("sccj")=request("qptjtz_sccj")
      rsadd("lydata")=request("qptjtz_lydata")
      rsadd("dqdata")=request("qptjtz_dqdata")
	  rsadd("cfdd")=request("qptjtz_cfdd")
      rsadd("bz")=request("qptjtz_bz") 
	  rsadd("tyqk")=request("qptjtz_tyqk")
	  rsadd("updata")=now()
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='qptjtz.asp';</Script>"
end sub

sub saveedit()    
      on error resume next
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from qptjtz where qptzid="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connjg,1,3
	  rsedit("sscj")=Trim(Request("qptjtz_sscj"))
      rsedit("bh")=request("qptjtz_bh")
      rsedit("bqname")=request("qptjtz_bqname") 
      rsedit("qptj")=Trim(request("qptjtz_qptj"))
      rsedit("qpyl")=request("qptjtz_qpyl")
      rsedit("qpcf")=request("qptjtz_qpcf")
      rsedit("ypbh")=request("qptjtz_ypbh")
	  rsedit("yt")=request("qptjtz_yt")
      rsedit("scdata")=request("qptjtz_scdata")
      rsedit("yxq")=request("qptjtz_yxq")
      rsedit("sccj")=request("qptjtz_sccj")
      rsedit("lydata")=request("qptjtz_lydata")
	  rsedit("dqdata")=request("qptjtz_dqdata")
      rsedit("cfdd")=request("qptjtz_cfdd")
	  rsedit("qpyq")=request("qptjtz_qpyq")   '
      rsedit("bz")=request("qptjtz_bz") 
	  rsedit("tyqk")=request("qptjtz_tyqk")
	  rsedit("updata")=now()
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from qptjtz_whjl where id="&id
  rsdel.open sqldel,connjg,1,3
set rsdel=nothing  
 
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from qptjtz where qptzid="&id
  rsdel.open sqldel,connjg,1,3
  dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub


sub edit()
  Dwt.out"<script type=""text/javascript"" src=""js/checkbh.js""></script>"&vbcrlf

   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from qptjtz where qptzid="&id
   rsedit.open sqledit,connjg,1,1
   dwt.out"<br><br><br><form method='post' action='qptjtz.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out"<table  border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out"<tr class='title'><td height='22' colspan='2'>"
   dwt.out"<div align='center'><strong>编辑气瓶管理台帐</strong></div></td>    </tr>"
	  dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>使用单位： </strong></td>"      
   dwt.out"<td  class='tdbg'>"
  if session("level")=0 then 
	dwt.out"<select name='qptjtz_sscj' size='1'>"
    dwt.out"<option >请选择使用单位</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out"</select></td></tr>  "  	 
  else 	 
    dwt.out"<input name='qptjtz_sscj' type='text' value='"&sscjh(rsedit("sscj"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
    dwt.out"<input name='qptjtz_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf
 end if 
 	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'>"
	 dwt.out"<strong>编&nbsp;&nbsp;号：</strong></td>"
	 dwt.out"<td  class='tdbg'><input name='qptjtz_bh' type='text' value='"&rsedit("bh")&"'></td>    </tr>   "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>标气名称：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_bqname' value='"&rsedit("bqname")&"' ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>体&nbsp;&nbsp;积：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qptj' value='"&rsedit("qptj")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>压&nbsp;&nbsp;力：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyl' value='"&rsedit("qpyl")&"'  ></td></tr> "
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>成份含量：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpcf' value='"&rsedit("qpcf")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>样品编号：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_ypbh' value='"&rsedit("ypbh")&"'  ></td></tr> "
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>余&nbsp;&nbsp;气：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_qpyq' value='"&rsedit("qpyq")&"'  ></td></tr> "

	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>用&nbsp;&nbsp;途：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_yt' value='"&rsedit("yt")&"'  ></td></tr> "
	 	 dwt.out"<tr class='tdbg'><td  align='right' class='tdbg'><strong>状态：</strong></td>"
	dwt.out"<td><select name='qptjtz_tyqk' size='1'>"
	sqlcj="SELECT distinct tyqk from qptjtz "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,connjg,1,1
		do while not rscj.eof
					dim tyqk
			select case rscj("tyqk")
			  case 1
				 tyqk="<span style='color:#006600'>在用</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>待换</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>退库</span>"
			end select

			dwt.out"<option value='"&rscj("tyqk")&"'"
			if cint(request("qptyqk"))=rscj("tyqk")  then dwt.out" selected"
			dwt.out ">"&tyqk&"</option>"& vbCrLf
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
    dwt.out"</select></td></tr>"


	 
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>定值日期</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
   dwt.out"<input name='qptjtz_scdata' style='WIDTH: 175px'  value='"&rsedit("scdata")&"'  onClick='new Calendar(0).show(this)' readOnly >"
	dwt.out"</td></tr>  "   
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>有效期：</strong></td> "
	Dwt.Out"<td width='80%' class='tdbg'>"	
	dwt.out outdatadict2 ("qptjtz_yxq","有效期",onnumb,rsedit("yxq"))

    dwt.out "</td></tr>"
	 
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>领用日期</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'>"
    dwt.out"<input name='qptjtz_lydata' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("lydata")&"'>"
	dwt.out"</td></tr>  "   
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>存放地点：</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_cfdd' value='"&rsedit("cfdd")&"' ></td></tr>  "  
	    
	 dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>供气厂家：</strong></td> "
	 dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_sccj' value='"&rsedit("sccj")&"' ></td></tr> "

	 dwt.out"<tr><td width='12%' align='right' class='tdbg'><strong>到期日期</strong></td>"      
	 Dwt.out "<td width='88%' class='tdbg'><input name='qptjtz_dqdata' type='text' id='input6' value='"&rsedit("dqdata")&"' onFocus='return addrdata()'/>&nbsp;<span >点击自动更新</span>"

	dwt.out"</td></tr>  "  
	 	
	dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out"<td width='88%' class='tdbg'><input type='text' name='qptjtz_bz' value='"&rsedit("bz")&"' ></td></tr>  "   
 
	 
	dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out"</table></form>"
	       rsedit.close
       set rsedit=nothing
	

end sub


sub main()
    url="qptjtz.asp"
    dim title,kk,qqadr
	sqlqptjtz="SELECT * from qptjtz "
	if keys<>"" then 
		sqlqptjtz=sqlqptjtz&" where bh like '%" &keys& "%' or bqname like'%" &keys& "%' or bqname like'%" &keys& "%'"
		title="-搜索 "&keys
		url="qptjtz.asp?keyword="&keys
	end if 
	if sscjid<>"" then
		sqlqptjtz=sqlqptjtz&" where sscj="&sscjid
		title="-"&sscjh(sscjid)
		url="qptjtz.asp?sscj="&sscjid
	end if 
	if qptyqk<>"" then
	if qptyqk=1 then
	   kk="在用"
	   else if qptyqk=2 then
	   kk="待换"
	   end if
	   kk="退库"
	   end if
		sqlqptjtz=sqlqptjtz&" where tyqk="&qptyqk
		title="-"&kk
		url="qptjtz.asp?tyqk="&qptyqk
	end if 
    if request("update")<>"" then 
		sqlqptjtz=sqlqptjtz&" ORDER BY updata desc"
    else
        sqlqptjtz=sqlqptjtz&" ORDER BY dqdata desc,SSCJ ASC"
	end if 

	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:5px;'>气瓶管理台帐</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	dim sql,rs,numb,totall,tyl,sqlty,sqlwjb,sqljb,sqlqx
	sql="select * from levelname where istq=false AND LEVELID<4"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		dwt.out "没有添加车间"
	else
	   do while not rs.eof
		sql="SELECT count(qptzid) FROM qptjtz WHERE sscj="&rs("levelid")
		sqlty="SELECT count(qptzid) FROM qptjtz WHERE sscj="&rs("levelid")&" and tyqk=1"'在用
		sqljb="SELECT count(qptzid) FROM qptjtz WHERE sscj="&rs("levelid")&" and tyqk=2"'待换

		sqlqx="SELECT count(qptzid) FROM qptjtz WHERE sscj="&rs("levelid")&" and tyqk=3"'退库
		numb=numb&sscjh_d(rs("levelid"))&":"&"<span style='color:#006600;'>"&connjg.Execute(sql)(0)&"/"&connjg.Execute(sqlty)(0)&"/"&connjg.Execute(sqljb)(0)&"</span>/<span style='color:#ff0000'>"&connjg.Execute(sqlqx)(0)&"&nbsp;&nbsp;"
	 
	rs.movenext
	loop
	end if 
	rs.close
	
	sql="SELECT count(qptzid) FROM qptjtz"
	totall= "<span style='color:#006600;'>"&connjg.Execute(sql)(0)&"</span>" 
	dwt.out "<div class='pre'><div align=left>数量/在用/待换/退库  "&numb&"合计:"&totall&"</div></div>"& vbCrLf

	call search()
	
	set rsqptjtz=server.createobject("adodb.recordset")
	rsqptjtz.open sqlqptjtz,connjg,1,1
	if rsqptjtz.eof and rsqptjtz.bof then 
	   message("未找到相关记录")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		dwt.out "     <td  class='x-td'><DIV class='x-grid-hd-text'>序号</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>使用单位</div></td>" & vbCrLf
		dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>气瓶编号</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>标气名称</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>体积</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>状态</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>操作</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>压力</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>成份含量</div></td>" & vbCrLf		
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>余气</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>定值日期</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>有效期</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>供气厂家</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>领用日期</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>到期日期</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>存放地点</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>用途</div></td>" & vbCrLf
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>退库日期</div></td>" & vbCrLf
'		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>更新日期</div></td>" & vbCrLf	
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>备注</div></td>" & vbCrLf		
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>" & vbCrLf
		dwt.out "    </tr>" & vbCrLf

	
			   record=rsqptjtz.recordcount
			   if Trim(Request("PgSz"))="" then
				   PgSz=20
			   ELSE 
				   PgSz=Trim(Request("PgSz"))
			   end if 
			   rsqptjtz.PageSize = Cint(PgSz) 
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
			   rsqptjtz.absolutePage = page
			   start=PgSz*Page-PgSz+1
			   rowCount = rsqptjtz.PageSize
			   do while not rsqptjtz.eof and rowcount>0
			   
			dim tyqk
			select case rsqptjtz("tyqk")
			  case 1
				 tyqk="<span style='color:#006600'>在用</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>待换</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>退库</span>"
			end select
			
if rsqptjtz("tyqk")="" then typk=""
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>" & vbCrLf

			dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh(rsqptjtz("sscj"))&"</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("bh")&"&nbsp;</div></td>" & vbCrLf

			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("bqname")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("qptj")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=qptjtz_whjl.asp?action=add1&qptjtzid="&rsqptjtz("qptzid")&">"&tyqk&"</a></div></td>" & vbCrLf
				qqadr=rsqptjtz("qptzid")
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><a href=qptjtz_whjl.asp?qptjtzid="&rsqptjtz("qptzid")&">更换</a></div></td>" & vbCrLf	
			if rsqptjtz("tyqk")=3 then 
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
				
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf	
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("tcdata")&"&nbsp;</div></td>" & vbCrLf
		
             else
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("qpyl")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("qpcf")&"&nbsp;</div></td>" & vbCrLf
			
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("qpyq")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("scdata")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("yxq")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("sccj")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"""
			if isnull(rsqptjtz("ghdata")) then
			dwt.out">"& rsqptjtz("lydata")&"&nbsp;"
			else
			dwt.out">"& rsqptjtz("ghdata")&"&nbsp;"
			end if
			dwt.out "</div></td>" & vbCrLf

			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("dqdata")&"&nbsp;</div></td>" & vbCrLf
				
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("cfdd")&"&nbsp;</div></td>" & vbCrLf
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("yt")&"&nbsp;</div></td>" & vbCrLf	
			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
			end if		

'			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("updata")&"&nbsp;</div></td>" & vbCrLf

			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsqptjtz("bz")&"&nbsp;</div></td>" & vbCrLf


			dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>" & vbCrLf
					call editdel(rsqptjtz("qptzid"),rsqptjtz("sscj"),"qptjtz.asp?action=edit&id=","qptjtz.asp?action=del&id=")
					
					dwt.out "</div></td></tr>" & vbCrLf
					 RowCount=RowCount-1
			  rsqptjtz.movenext
			  loop
			  
			  
			dwt.out "</table>" & vbCrLf
		    if sscjid<>"" or keys<>"" then 
			  call showpage(page,url,total,record,PgSz)
			else
			  call showpage1(page,url,total,record,PgSz)
			end if  
		   end if
		   rsqptjtz.close
		   set rsqptjtz=nothing
			connjg.close
			set connjg=nothing
end sub





dwt.out "</body></html>"



sub search()
	dim sqlcj,rscj,sqlgh,rsgh
	dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	dwt.out "<div class='x-toolbar'><div align=left>" & vbCrLf
	dwt.out "<form method='Get' name='SearchForm' action='qptjtz.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then  dwt.out "<a href='qptjtz.asp?action=add'>添加气瓶</a>"
	dwt.out "&nbsp;&nbsp;<a href='qptjtz.asp?update=updata'>查看最近更新</a>"
	dwt.out "  <input type='text' name='keyword'  size='20' maxlength='50' "

	if keys<>"" then 
	 dwt.out "value='"&keys&"'"
    	dwt.out ">" & vbCrLf
    else
	 dwt.out "value='输入搜索字'"
	 	dwt.out " onblur=""if(this.value==''){this.value='输入搜索字'}"" onfocus=""this.value=''"">" & vbCrLf
	end if                 
	dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
	dwt.out "&nbsp;&nbsp;<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			dwt.out"<option value='qptjtz.asp?sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid")  then dwt.out" selected"
			dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
		
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf
	dwt.out "&nbsp;&nbsp;<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		dwt.out "	       <option value=''>请选择投运情况</option>" & vbCrLf
	sqlcj="SELECT distinct tyqk from qptjtz "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,connjg,1,1
		do while not rscj.eof
					dim tyqk
			select case rscj("tyqk")
			  case 1
				 tyqk="<span style='color:#006600'>在用</span>"
			  case 2 
				tyqk="<span style='color:#0000ff'>待换</span>"
			  case 3 
				tyqk="<span style='color:#ff0000'>退库</span>"
			end select

			dwt.out"<option value='qptjtz.asp?styqk="&rscj("tyqk")&"'"
			if cint(request("qptyqk"))=rscj("tyqk")  then dwt.out" selected"
			dwt.out ">"&tyqk&"</option>"& vbCrLf
			rscj.movenext
		loop
		rscj.close
		set rscj=nothing
		dwt.out "     </select>	" & vbCrLf

		dwt.out "</form></div></div>" & vbCrLf

end sub





Call Closeconn
%>