<!--#include file="conn.asp"-->
<%
Dim dwt
Set dwt= New dwt_Class	
Class dwt_Class
	Public Function out(s) 
		response.write s
	End Function 



End Class

Dim Action, FoundErr, ErrMsg, ComeUrl,total1
Dim strInstallDir


'****************************
'文件结尾版权
'***********************888
sub footer()
response.write "<br><br>"
response.write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""100%"" class=""border"" align=center>"
response.write "<tr align=""center"">"
response.write "<td height=25 class=""topbg""><span class=""Glow"">设备管理系统 All Rights Reserved.</span>"
response.write "</tr></table></body></html>"
end sub



'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceBadChar(strChar)
    strChar=REPLACE(STRCHAR,"'","")
    ReplaceBadChar = strChar
End Function

Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Function PE_CDbl(ByVal str1)
    If IsNumeric(str1) Then
        PE_CDbl = CDbl(str1)
    Else
        PE_CDbl = 0
    End If
End Function








'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
'URL中带？的
'*******************************************
sub showpage(page,url,total,record,pgsz)
   response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'><tr class='tdbg'><td><div align=""center"">"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  if Total =1 then 
    response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  else
   response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div></td></tr></table>"
end sub


'百分之80表格显示
sub showpage_80(page,url,total,record,pgsz)
   response.write "<table width='80%' align='center'  border='0' cellpadding='0' cellspacing='0' class='border'><tr class='tdbg'><td><div align=""center"">"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  if Total =1 then 
    response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  else
   response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div></td></tr></table>"
end sub


'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
 'url中不带？
'*******************************************
sub showpage1(page,url,total,record,pgsz)
   response.write "<table width='100%' border='0' cellpadding='0' cellspacing='0' class='border'><tr class='tdbg'><td><div align=""center"">"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"?page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"?pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"?pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"?pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
   if Total =1 then 
     response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">条"
   else
     response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">条"
   end if 
   if Total=1 then 
       response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
       response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div></td></tr></table>"
end sub



'1维修一车间，2维修二车间，3维修三车间，4维修四车间，5综合车间，6计量科
Function sscjh(sscj)
    dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    sscjh=rscj("levelname")
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
end Function 

'用于短的车间显示
Function sscjh_d(sscj)
       dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    sscjh_d=replace(replace(replace(rscj("levelname"),"修",""),"车间",""),"科","")
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
end Function 

 '29日新加
'用于编辑新增装置显示
function formgh(ghid,sscj)
	dim sqlgh,rsgh
	
'
if isnull(sscj) then sscj=0
if isnull(ghid) then ghid=0
if sscj=0 then 
	sqlgh="SELECT * from ghname"
else
sqlgh="SELECT * from ghname where sscj="&sscj
end if 		
set rsgh=server.createobject("adodb.recordset")
		rsgh.open sqlgh,conn,1,1
		if rsgh.eof then 
		gh="未编辑"
	else
		response.write"<option value='0'"
		if ghid=0 then response.write " selected" 
		response.write">请选择装置</option>"
		do while not rsgh.eof
			response.write"<option value='"&rsgh("ghid")&"' "
			if ghid=rsgh("ghid") then response.write "selected"
			response.write">"&rsgh("gh_name")&"</option>"  & vbCrLf   
		rsgh.movenext
	loop
end if 
	rsgh.close
	set rsgh=nothing

end function
 '29日新加
'取装置工号名称
Function gh(ghid)
       dim sqlgh,rsgh
	if isnull(ghid) then ghid=0
	sqlgh="SELECT * from ghname where ghid="&ghid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    if rsgh.eof then 
	  gh="未编辑"
	else
	    gh=rsgh("gh_name")
end if 
	rsgh.close
	set rsgh=nothing
end Function 
 '29日新加
'取分级的星数
Function fj(fjnumb)
       dim fj_i
	if isnull(fjnumb) or fjnumb=0 then 
	  fj="未分级"
	else
		for fj_i=1 to fjnumb
		fj=fj&"*"
		next
	  
	end if 
end Function 





'热动班1,供水班2,合成一班3,合成二班4,气压班5,复肥班6,硝铵班7,硝酸班8
Function ssbzh(ssbz)
            dim sqlbz,rsbz
	  sqlbz="SELECT * from bzname where id= "&ssbz
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    do while not rsbz.eof
       	'response.write"<option value='"&rsbz("levelid")&"'>"&rsbz("levelname")&"</option>"& vbCrLf
	    ssbzh=rsbz("bzname")
		rsbz.movenext
	loop
	rsbz.close
	set rsbz=nothing
end Function


'选项（编辑、删除）
sub editdel(id,sscj,editurl,delurl)
 if session("level")=sscj or session("level")=0 then 
    response.write "<a href="&editurl&id&">编辑</a>&nbsp;"
	response.write "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 else
    response.write "&nbsp;"
 end if 
end sub

'选项（编、删）
sub editdel_d(id,sscj,editurl,delurl)
 if session("level")=sscj or session("level")=0 then 
    response.write "<a href="&editurl&id&">编</a>&nbsp;"
	response.write "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删</a>"
 else
    response.write "&nbsp;"
 end if 
end sub
 
 
 '高亮显示搜索关键字
function searchH(body,key)
dim inum'第一次出现的位置
dim leftbody '截取BODY中的字符串到INUM
dim lenkey   'KEY的字符串长度
dim lenbody 'BODY的字符串长度
dim rightbody '截取BODY中的字符串到KEY的尾部
dim midikey  '将KEY高亮显示为蓝色
key=UCase(cstr(key))
body=UCase(cstr(body))
inum=InStr(body,key)
lenkey=len(key)
lenbody=len(body)
on error resume next
leftbody=left(body,inum-1)
rightbody=right(body,lenbody-inum-lenkey+1)
midikey="<font color='#0000FF'><strong>"&key&"</strong></font>"
searchH=leftbody&midikey&rightbody
end function 

'在增加表单中显示选择车间的表单
function formsscj()
   if session("level")=0 then 
	response.write"<select name='sscj' size='1'>"
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
    response.write"</select>"  	 
   else 	 
     response.write"<input name='sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
     response.write"<input name='sscj' type='hidden' value="&session("level")&">"& vbCrLf
  end if 

end function


'在增加表单中显示选择车间和班组的表单
function formsscjbz()
 dim rscj,sqlcj,rsbz,sqlbz
 if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	response.write"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    response.write"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select>"  	 & vbCrLf
    response.write "<select name='ssbz' size='1' >" & vbCrLf
    response.write "<option  selected>选择班组分类</option>" & vbCrLf
    response.write "</select></td></tr>  "  & vbCrLf
    response.write "<script><!--" & vbCrLf
    response.write "var groups=document.form1.sscj.options.length" & vbCrLf
    response.write "var group=new Array(groups)" & vbCrLf
    response.write "for (i=0; i<groups; i++)" & vbCrLf
    response.write "group[i]=new Array()" & vbCrLf
    response.write "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0	
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   response.write "group["&rscj("levelid")&"][0]=new Option(""未添加班组"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'response.write"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   response.write"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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
    response.write "var temp=document.form1.ssbz" & vbCrLf
    response.write "function redirect(x){" & vbCrLf
    response.write "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    response.write "temp.options[m]=null" & vbCrLf
    response.write "for (i=0;i<group[x].length;i++){" & vbCrLf
    response.write "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    response.write "}" & vbCrLf
    response.write "temp.options[0].selected=true" & vbCrLf
    response.write "}//--></script>" & vbCrLf
  else 	 
   response.write"<input name='sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
   response.write"<input name='sscj' type='hidden' value="&session("level")&">"& vbCrLf
   sqlbz="SELECT * from bzname where sscj="&session("level")
   set rsbz=server.createobject("adodb.recordset")
   rsbz.open sqlbz,conn,1,1
   response.write"<select name='ssbz' size='1'>"
   
   if rsbz.eof and rsbz.bof then 
   	  response.write"<option value='0'>未添加班组</option>"
   else   
	  'response.write"<option value='0'>车间</option>"
      do while not rsbz.eof
	     response.write"<option value='"&rsbz("id")&"'>"&rsbz("bzname")&"</option>"
	  rsbz.movenext
      loop
	  end if 
	 response.Write"</select>" 
  rsbz.close
  set rsbz=nothing
 end if 
end function


'取当前网页URL
Function GetUrl() 
	'On Error Resume Next 
	Dim strtemp 
	If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
	 strtemp = "http://"
	Else 
	 strtemp = "https://"
	End If 
	strtemp = strtemp & Request.ServerVariables("SERVER_NAME") 
	If Request.ServerVariables("SERVER_PORT") <> 80 Then 
	 strtemp = strtemp & ":" & Request.ServerVariables("SERVER_PORT") 
	end if
	strtemp = strtemp & Request.ServerVariables("URL") 
	If Trim(Request.QueryString) <> "" Then 
	 strtemp = strtemp & "?" & Trim(Request.QueryString) 
	end if
	'判断URL中是否有分页函数，有则去掉
	if InStr(strtemp,"pgsz")<>0 then
		urllen=InStr(strtemp,"pgsz")
		strtemp=left(strtemp,urllen-2)
	end if  
	GetUrl = strtemp 
End Function


function message(text)
response.write "<br><br><br><table width=""351"" height=""185"" border='0' align='center' cellpadding='2' cellspacing='1' class='border' >" & vbCrLf
response.write "  <tr  class='title'>" & vbCrLf
response.write "    <td height=""30"" colspan=""3""><div align=center>系统提示信息</div></td>" & vbCrLf
response.write "  </tr>" & vbCrLf
response.write "  <tr>" & vbCrLf
response.write "    <td><div align=""center""><img src=""/images/3.gif"" width=""60"" height=""60""></div></td>" & vbCrLf
response.write "    <td>&nbsp;</td>" & vbCrLf
response.write "    <td>"&text&"</td>" & vbCrLf
response.write "  </tr>" & vbCrLf
response.write "</table>" & vbCrLf

end function
%>