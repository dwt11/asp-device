    
    <%
	'获取分类名称
	Function getclassname(classid)
       dim sqlgh,rsgh
	if isnull(classid) then classid=0
	sqlgh="SELECT * from dgtzl_index_gh where id="&classid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conndgt,1,1
    if rsgh.eof then 
	  getclassname="未编辑"
	else
	    'getclassname="<a href=showlist.asp?classid="&classid&">"&rsgh("class_name")&"</a>"
		getclassname=rsgh("class_name")
end if 
	rsgh.close
	set rsgh=nothing
end Function 


	'获取标题
	Function gettitlename(bodyid)
       dim sqlgh,rsgh
	if isnull(bodyid) then bodyid=0
	sqlgh="SELECT * from dgtzl_body where id="&bodyid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conndgt,1,1
    if rsgh.eof then 
	  gettitlename="未编辑"
	else
	    gettitlename=rsgh("news_title")
	end if 
		rsgh.close
		set rsgh=nothing
end Function 






	'循环获取上级分类名称于当前显示
	Function gettclassname(classid)
       dim sqlgh,rsgh
	if classid="" then 
	else
		sqlgh="SELECT * from dgtzl_index_gh where id="&classid
		set rsgh=server.createobject("adodb.recordset")
		rsgh.open sqlgh,conndgt,1,1
		if rsgh.eof then 
		    gettclassname="> 未编辑"
		else
		
		  if rsgh("index")<>0 then 
			  sqlgh1="SELECT * from dgtzl_index_gh where id="&rsgh("index")
			  set rsgh1=server.createobject("adodb.recordset")
			  rsgh1.open sqlgh1,conndgt,1,1
			  if rsgh1.eof then 
				  gettclassname="> 未编辑"
			  else
					gettclassname=" > <a href=index.asp?classid="&rsgh1("id")&">"&rsgh1("class_name")&"</a>"
			  end if 
		  end if 
		
		'gettclassname="> "&gettclassname&" > "&rsgh("class_name")
		
		
		
		
		end if 
			rsgh.close
			set rsgh=nothing
	end if 
	end Function 
	
	
	
	'获取短日期 月-日 
	Function getsdate(newdate)
	getsdate=month(newdate)&"-"&day(newdate)
	end Function 
	
	'获取年-月-日 
	Function gettdate(newdate)
	gettdate=year(newdate)&"-"&month(newdate)&"-"&day(newdate)
	end Function 
	
	'格式 化内容中图片大小 
	  function imgCode(strContent)
	  dim re
	  Set re=new RegExp
	  re.IgnoreCase =true
	  re.Global=True
			  
	  re.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
	  strContent=re.replace(strContent,"<div align=center><img SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>550)this.width=333""></div>")
	  
	  
	  set re=Nothing
	  imgCode=strContent
	  end function	
	
	'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
'URL中带？的
'*******************************************
sub newshowpage(page,url,total,record,pgsz)
   response.write "<div class='page'>"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1>首页</a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&">上一页</a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&">下一页</a> <a href="&url&"&pgsz="&pgsz&"&page="&total&">尾页</a>"
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
   response.write "</div>"
end sub


	%>