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
<%response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title> 计量管理管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dim sqlcj,rscj,i,ii,sqlbz,rsbz,sql,rs
if Request("action")="zjinfo" then call zjinfo
if request("action")="complete" then call complete
if request("action")="completesave" then call completesave


sub zjinfo()
'************************************算法

'在ZJTZ表中遍历所有需周检的表，
'如果检定周期是0或1（停用OR不周检）跳过
'在ZJINFO表中从最新ID开始查找遍历过来的ZJTZID的表
'如果ZJYEAR＝提交过来的，则输出计划 月份、日期、结果为ZJINFO表中存的数据
'如果ZJYEAR＜＞提交过来的，则输出计划 月份为提交过来的数据，并且鉴定日期和结果为空白

'此段代码有待改
'***********************************88
dim zjinfoor    '用于判断是否找到相应的周检信息
	zjinfoor=0
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from zjtz where sscj="&cint(request("sscj"))&" and ssbz="&cint(request("ssbz"))&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
      dim text
	 zjmonth=request("zjmonth")
	 if cint(zjmonth)=0 then 
	  zjmonthname="大修"
	  text="未找到 "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonthname&"   周检情况"
	 else
	  text="未找到 "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonth&"月   周检情况"
	 end if  
	  call message(text)
   else
      response.write "<table height=50 width=""100%"" border=""0"" align=""center"" cellpadding=""0""><tr><td height=40><font size=""5""><div align=center>"
	  if cint(request("zjmonth"))=0 then
	     zjmonth="大修"
	     response.write sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonthname&"   周检情况"
	  else
   	     zjmonth=cint(request("zjmonth"))
		 response.write sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonth&"月  周检情况"
	  end if    
	  response.write "</div></font></td></tr></table>"
	  response.write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
      response.write "<tr class=""title"">"  & vbCrLf
      response.write "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>序号</strong></div></td>" & vbCrLf
      response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>车间</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>类型</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>位号</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>规格型号</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>测量范围</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>鉴定周期</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>计划鉴定月份</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>鉴定日期</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>鉴定结果</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备注</strong></div></td>" & vbCrLf
      response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>&nbsp;</strong></div></td>" & vbCrLf
      response.write "    </tr>" & vbCrLf
      do while not rszjtz.eof
          dim jdzq  '检定周期判断
		  dim jdyear '检定周期换算为年
		  jdzq=rszjtz("jdzq")
		  if jdzq=0 then 
			  'response.write "<td><font color=#ff0000><div align=center>停用</div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
		  else
		      if jdzq=1 then 
    		      'response.write "<td><font color=#ff0000><div align=center>不周检</div></font></td><td>&nbsp;</td><td>&nbsp;</td>"
			  else
				  jdyear=jdzq/12
		          sqlscdate="SELECT * from zjinfo where zjtzid="&rszjtz("id")&" ORDER BY id DESC"
				  'zjyear="&request("zjyear")-jdyear&" and zjmonth="&request("zjmonth")
                  set rsscdate=server.createobject("adodb.recordset")
                  rsscdate.open sqlscdate,connzj,1,1
                  if rsscdate.eof and rsscdate.bof then 
                       response.write "<td><div align=center>未找到内容,请在周检台账中添加此表的初次周检日期</div></td></tr>" 
                  else
					   if rsscdate("zjyear")=cint(request("zjyear")) and rsscdate("zjmonth")=cint(request("zjmonth"))  then
                              response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                              response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rszjtz("id")&"</div></td>" & vbCrLf
                              response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_D(rszjtz("sscj"))&ssbzh(rszjtz("ssbz"))&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&zjclass(rszjtz("class"))&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("wh")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("ggxh")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("clfw")&"&nbsp;</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("jdzq")&"&nbsp;</div></td>" & vbCrLf
				              if rsscdate("zjmonth")=0 then 
							     zjmonthname="大修"
							  else
							     zjmonthname=rsscdate("zjmonth")
							  end if  
							  response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjyear")&"-"&zjmonthname&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjday")&"</div></td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsscdate("zjinfo")&"</div></td>" & vbCrLf
		                      response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("bz")&"&nbsp;</td>" & vbCrLf
                              response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center>&nbsp;</div>" & vbCrLf
                              response.write "</td></tr>" & vbCrLf
						zjinfoor=1
						else 
							  if rsscdate("zjyear")=cint(request("zjyear"))-jdyear and rsscdate("zjmonth")=cint(request("zjmonth"))  then
                                     response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                                     response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rszjtz("id")&"</div></td>" & vbCrLf
                                     response.write "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&sscjh_D(rszjtz("sscj"))&ssbzh(rszjtz("ssbz"))&"</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&zjclass(rszjtz("class"))&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("wh")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("ggxh")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("clfw")&"&nbsp;</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rszjtz("jdzq")&"&nbsp;</div></td>" & vbCrLf
				                     if request("zjmonth")=0 then 
							            zjmonthname="大修"
							         else
							            zjmonthname=request("zjmonth")
							         end if  
				                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&request("zjyear")&"-"&zjmonthname&"</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">&nbsp;</div></td>" & vbCrLf
		                             response.write "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rszjtz("bz")&"&nbsp;</td>" & vbCrLf
                                     response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=center><a href=zjqk.asp?action=complete&id="&rszjtz("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&">完成</aS></div>" & vbCrLf
                                     response.write "</td></tr>" & vbCrLf
                              'else
							         'response.write "<td><div align=center>未找到相关内容</div></td></tr>" 
							  						zjinfoor=1
							  end if 
						end if 	  
				end if 
			    rsscdate.close
		     end if 
	     end if   
    rszjtz.movenext
 
 loop
    response.write "</table>" & vbCrLf
 
 '判断上面的循环是否找到相关内容，并调用消息提示框，如果找到相关信息则输出，导出和返回
  if zjinfoor=0 then 
   if cint(request("zjmonth"))=0 then 
	  zjmonthname="大修"
	  text="未找到 "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonthname&"   周检情况"
	 else
	  text="未找到 "&sscjh(request("sscj"))&" "&ssbzh(request("ssbz"))&" "&request("zjyear")&"年"&zjmonth&"月   周检情况"
	 end if  
	  call message(text)
   else
   			response.write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  class='border'><tr class='tdbg'><td><div align=right>"
			response.write "<input type='button' name='Submit'  onclick=""window.location.href='tocsv.asp?action=zjtz&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&"&titlename=周检台账'"" value='导出上面内容到EXCEL'>"
			
			response.write "</div></td></tr></table>"

   
   end if 	  
 
 
   end if
   rszjtz.close
   set rszjtz=nothing
end sub
response.write "</body></html>"


'用于保存本月周检完成后所添的周检结果
sub complete()
   dim sqlzjtz,rszjtz,rsscdate,sqlscdate,zjmonth,zjmonthname
   sqlzjtz="SELECT * from zjtz where id="&request("id")&" ORDER BY id DESC"
   set rszjtz=server.createobject("adodb.recordset")
   rszjtz.open sqlzjtz,connzj,1,1
   if rszjtz.eof and rszjtz.bof then 
        message("未知错误")
   else
   response.write"<br><br><br><form method='post' action='zjqk.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   response.write"<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   response.write"<tr class='title'><td height='22' colspan='2'>"
   response.write"<div align='center'><strong>周检结果添写</strong></div></td>    </tr>"
   response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"      
   response.write"<td width='88%' class='tdbg'><input disabled='disabled'  type='text' value='"&sscjh(rszjtz("sscj"))&"' size=10>&nbsp;<input disabled='disabled'  type='text' value='"&ssbzh(rszjtz("ssbz"))&"' size=8></td></tr>"& vbCrLf
	
	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>位&nbsp;&nbsp;号：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("wh")&"></td>    </tr>   "
	 
	 
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>类&nbsp;&nbsp;型：</strong></td> "
	response.write"<td><input disabled='disabled' type='text' value="&zjclass(rszjtz("class"))&"></td></tr>"
	 
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>规格型号：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("ggxh")&"></td>    </tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>测量范围：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("clfw")&"></td>    </tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>鉴定周期：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input disabled='disabled' type='text' value="&rszjtz("jdzq")&"></td></tr>"
	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>周检月份：</strong></td>"
    'dim jdyear,zjyear
	'jdyear=rszjtz("jdzq")/12
    'sqlscdate="SELECT * from zjinfo where zjtzid="&rszjtz("id")&" ORDER BY id DESC"
    'set rsscdate=server.createobject("adodb.recordset")
    'rsscdate.open sqlscdate,connzj,1,1
    'if rsscdate.eof and rsscdate.bof then 
     '   response.write "<td><div align=center>未找到内容,请在周检台账中添加此表的初次周检日期</div></td></tr>" 
     'else
	 'zjyear=rsscdate("zjyear")+jdyear
	 zjmonthname=request("zjmonth")
	 if zjmonthname=0 then zjmonthname="大修"
	 response.write"<td width='80%' class='tdbg'><input disabled='disabled' type='text' value="&request("zjyear")&"-"&zjmonthname&"></td>    </tr>   "
    'end if 
	
	response.write"<input type='hidden' name=""zjyear"" value='"&request("zjyear")&"'>"
	response.write"<input type='hidden' name=""zjmonth"" value='"&request("zjmonth")&"'>"
	response.write"<input type='hidden' name=""sscj"" value='"&request("sscj")&"'>"
	response.write"<input type='hidden' name=""ssbz"" value='"&request("ssbz")&"'>"

	response.write"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>周检日期：</strong></td>"
	 response.write"<td width='80%' class='tdbg'>"
	 response.write"<select name=zjday>"
	 dim i
	 for i=1 to 31
	  response.write "<option value='"&i&"'"& vbCrLf
	  if i=day(now()) then response.write "selected"
	  response.write">"&i&"</option>"& vbCrLf
	 next
	 response.write"</select></td></tr>   "
    response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>鉴定结果：</strong></td>"
	 response.write"<td width='88%' class='tdbg'><input name='zjinfo' type='text'></td>    </tr>   "

	response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    response.write"<td width='88%' class='tdbg'><input type='text' name='bz'></td></tr>  "   

	response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	response.write"<input name='action' type='hidden' id='action' value='completesave'> <input type='hidden' name='id' value='"&request("id")&"'>     <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back(-1)"" style='cursor:hand;'></td>  </tr>"
	response.write"</table></form>"
    'response.write request("sscj")&&
   end if 
end sub



sub completesave()
      dim rsadd,sqladd
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from zjinfo" 
      rsadd.open sqladd,connzj,1,3
      rsadd.addnew
      rsadd("zjtzid")=Trim(Request("id"))
      rsadd("zjyear")=cint(Request("zjyear"))
	  rsadd("zjmonth")=cint(request("zjmonth"))
      rsadd("zjday")=request("zjday")
      rsadd("bz")=request("bz")
      rsadd("zjinfo")=request("zjinfo")
	  rsadd.update
rsadd.close
	  response.write"<Script Language=Javascript>location.href='zjqk.asp?action=zjinfo&sscj="&request("sscj")&"&ssbz="&request("ssbz")&"&zjyear="&request("zjyear")&"&zjmonth="&request("zjmonth")&"';</Script>"

end sub

'用于分类名称显示
Function zjclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from class where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	'do while not rsname.eof
	else
	    zjclass=rsname("name")
		'rsname.movenext
	'loop
	end if 
	rsname.close
	set rsname=nothing
end Function

Call Closeconn
%>