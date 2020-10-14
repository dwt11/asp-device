<%@language=vbscript codepage=65001 %>

<%
Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->


<%
dim sqllsda,rslsda,title,record,pgsz,total,page,start,rowcount,xh,url,ii,zxzz
dim rsadd,sqladd,lsdaid,rsedit,sqledit,scontent,rsdel,sqldel,tyzk,id
url="lsda.asp"


if Request("action")="add" then call add
if Request("action")="saveadd" then call saveadd
if request("action")="edit" then call edit
if request("action")="saveedit" then call saveedit
if request("action")="del" then call del

sub add()
dim rscj,sqlcj
   response.write"<br><br><br><form method='post' action='lsda_action.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' bgcolor=""#FFFFFF""  cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>添加联锁档案</strong></div></td>    </tr>"
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
    response.write"<td width='80%' class='tdbg'>"
  if session("level")=0 then 
	response.write"<select name='lsda_sscj' size='1'>"
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
     response.write"<input name='lsda_sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' ></td></tr>"& vbCrLf
      response.write"<input name='lsda_sscj' type='hidden' value="&session("level")&"></td></tr>"& vbCrLf

 end if 

	 
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 response.write"<strong>位&nbsp;&nbsp;号：</strong></td>"
	 response.write"<td width='80%' class='tdbg'><input name='lsda_wh' type='text'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>用&nbsp;&nbsp;途：</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_yt' ></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>一次元件名称：</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_ycjname'></td></tr> "
	 
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量单位：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_cldw'></td></tr>  "   
   
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量范围：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_clfw'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>联锁值L：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_lsl'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>联锁值H：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_lsh'></td></tr>  "   
   
    response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>投运状况：</strong></td>"
	response.write"<td><select name='lsda_tyzk' size='1'>"
	response.write"<option value='1'>投运</option>"
    response.write"<option value='0'>旁路</option>"
    response.write"</select></td></tr>"
	
    response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>执行装置：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_zxzz'></td></tr>  "   
	 
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<a href=""#"" class=""lbAction"" rel=""deactivate""><button>取消</button></a></td>  </tr>"
	response.write"</table></form>"
end sub	

sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from lsda" 
      rsadd.open sqladd,connjg,1,3
      rsadd.addnew
      rsadd("sscj")=Trim(Request("lsda_sscj"))
      rsadd("wh")=request("lsda_wh")
      rsadd("yt")=Trim(request("lsda_yt"))
      rsadd("ycjname")=request("lsda_ycjname")
      rsadd("cldw")=request("lsda_cldw")
      rsadd("clfw")=request("lsda_clfw")
      rsadd("lsl")=request("lsda_lsl")
      rsadd("lsh")=request("lsda_lsh")
      rsadd("tyzk")=request("lsda_tyzk")
      rsadd("zxzz")=request("lsda_zxzz")
      rsadd("bz")=request("lsda_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  response.write"<Script Language=Javascript>location.href='lsda.asp';</Script>"
end sub


sub saveedit()    
	  '保存
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from lsda where lsdaid="&ReplaceBadChar(Trim(request("ID")))
      rsedit.open sqledit,connjg,1,3
      rsedit("sscj")=Trim(Request("lsda_sscj"))
      rsedit("wh")=request("lsda_wh")
      rsedit("yt")=Trim(request("lsda_yt"))
      rsedit("ycjname")=request("lsda_ycjname")
      rsedit("cldw")=request("lsda_cldw")
      rsedit("clfw")=request("lsda_clfw")
      rsedit("lsl")=request("lsda_lsl")
      rsedit("lsh")=request("lsda_lsh")
      rsedit("tyzk")=request("lsda_tyzk")
      rsedit("zxzz")=request("lsda_zxzz")
      rsedit("bz")=request("lsda_bz")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  response.write"<Script Language=Javascript>history.go(-1)</Script>"
end sub

sub del()
  lsdaid=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from lsda where lsdaid="&lsdaid
  rsdel.open sqldel,connjg,1,3
  response.write"<Script Language=Javascript>history.go(-1)</Script>"
  
set rsdel=nothing  

end sub


sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from lsda where lsdaid="&id
   rsedit.open sqledit,connjg,1,1
   response.write"<br><br><br><form method='post' action='lsda_action.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='100%' border='0' align='center' cellpadding='2' bgcolor=""#FFFFFF"" cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>编辑联锁档案</strong></div></td>    </tr>"
     
     response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     response.write"<td width='80%' class='tdbg'><input name='lsda_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     response.write"<input name='lsda_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

	 
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'>"
	 response.write"<strong>位&nbsp;&nbsp;号：</strong></td>"
	 response.write"<td width='80%' class='tdbg'><input name='lsda_wh' type='text' value='"&rsedit("wh")&"'></td>    </tr>   "
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>用&nbsp;&nbsp;途：</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_yt'  value='"&rsedit("yt")&"'></td></tr> "
	 
	 response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>一次元件名称：</strong></td> "
	 response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_ycjname' value='"&rsedit("ycjname")&"'></td></tr> "
	 
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量单位：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_cldw' value='"&rsedit("cldw")&"'></td></tr>  "   
   
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>测量范围：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_clfw' value='"&rsedit("clfw")&"'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>联锁值L：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_lsl' value='"&rsedit("lsl")&"'></td></tr>  "   
   	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>联锁值H：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_lsh' value='"&rsedit("lsh")&"'></td></tr>  "   
   
    response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>投运状况：</strong></td>"
	response.write"<td><select name='lsda_tyzk' size='1'>"
	response.write"<option value='1'"
	if rsedit("tyzk")=1 then response.write"selected"
	response.write">投运</option>"
    response.write"<option value='0'"
	if rsedit("tyzk")=0 then response.write"selected"
	response.write">旁路</option>"
    response.write"</select></td></tr>"
	
    response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>执行装置：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_zxzz' value='"&rsedit("zxzz")&"'></td></tr>  "   
	 
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='80%' class='tdbg'><input type='text' name='lsda_bz' value='"&rsedit("bz")&"'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<a href=""#"" class=""lbAction"" rel=""deactivate""><button>取消</button></a></td>  </tr>"
	response.write"</table></form>"
       rsedit.close
       set rsedit=nothing
	
end sub


Call Closeconn
%>