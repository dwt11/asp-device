<%
class dwt_page
'分页类

public  dwtconn    	 'connection数据库连接对象
public  dwtrs      	 'recordset对象
public  dwttable   	 '要分页的table数据库表
public  dwtpagesize      '分页的pagesize大小
public  dwtpage          'URL后面的page参数值
public  dwtrssql         '分页的SQL语句
public  dwttourl         '链接的URL
public  dwtrecordcount   '数据库记录数
public  dwtpagecount     '总页码数
public  dwtpagestyle     '分页的样式，取值为1或2，样式1适合页数较多的，样式2适合页数较少的
public  dwtsql    '数据检索时的条件,引用时必须是" where 表达式 and "前后空格都要有

Private Sub Class_Initialize '类的初始化

	dwtpagesize=10 '默认分页大小
	dwtpagestyle=1 '默认分页显示样式
	
	if isempty(request.querystring("page")) <> true then '检证page后面的参数
	
			if isnumeric(request.querystring("page"))=false or request.querystring("page")="" then
			
					dwtpage=1

			else
					if   request.querystring("page")<=0 then
					
						dwtpage=1
						
					else
					
						dwtpage=request.querystring("page")
					
					end if
			
			end if
			
		else
		
			dwtpage=1
			
	end if

End Sub

Public Property Let getconn(zwconn) '得到数据库连接的connection对象
	
	set dwtconn=zwconn
	
End Property 

public sub dwt_set '根据提供的参数进行设置

	dwtrecordcount=dwtconn.execute("select count(id) from " & dwttable&replace(dwtsql,"and",""))(0) '数据库总记录数
	
	if dwtrecordcount mod dwtpagesize=0 then '获得最大页码
	
		dwtpagecount=int(dwtrecordcount/dwtpagesize)
		
	else
	
		dwtpagecount=int(dwtrecordcount/dwtpagesize)+1
		
	end if
	
	if clng(dwtpage) > clng(dwtpagecount) then '如果page后面的页码数大于最大页码，取最大页码
	
		dwtpage=dwtpagecount
		
	end if
	
	dwtrssql="select top " & dwtpagesize  & " * from  " & dwttable '动态组建分页的SQL语句
	dwtrssql=dwtrssql & " where "&replace(dwtsql,"where","")&" ID<=( select min(ID) from (select top " & cint(dwtpagesize)*cint(dwtpage-1)+1
	dwtrssql=dwtrssql & " ID from " & dwttable & " order by id desc) as tabview) order by  id desc"
	
	set dwtrs=dwtconn.execute(dwtrssql) '返回一个只读向前的recordset对象
	
end sub


private function topage(n) '翻页

	dwttourl="?page=" & n 
	
	for each str in request.querystring 'URL后面的所有参数
	
			if lcase(str) <> "page" then
			
				dwttourl=dwttourl & "&" & str & "=" & request.querystring(str)
				
			end if
	
	next
	
	topage=dwttourl
	
end function

public sub dwtshowpage '显示分页

	dim showpagehtml
	
	select case dwtpagestyle
		
		case 1 '分页样式1
		if cint(dwtpage)<=1 then
		
			showpagehtml="首页&nbsp;&nbsp;上一页&nbsp;&nbsp;<a href=""" & topage(dwtpage+1) & """>"
			showpagehtml=showpagehtml & "下一页</a>&nbsp;&nbsp;<a href=""" & topage(dwtpagecount) & """>尾页</a>"
			
		elseif cint(dwtpage)>=dwtpagecount then
		
			showpagehtml="<a href=""" & topage(1) & """>首页</a>&nbsp;&nbsp;<a href=""" & topage(dwtpage-1) & """>上一页</a>"
			showpagehtml=showpagehtml & "&nbsp;&nbsp;下一页&nbsp;&nbsp;尾页"
			
		else
		
			showpagehtml="<a href=""" & topage(1) & """>首页</a>&nbsp;&nbsp;<a href=""" & topage(dwtpage-1) & """>上一页</a>&nbsp;&nbsp;"
			showpagehtml=showpagehtml & "<a href=""" & topage(dwtpage+1) & """>下一页</a>&nbsp;&nbsp;<a href=""" & topage(dwtpagecount) & """>尾页</a>"
			
		end if
		
		response.write "<script>"
		response.write "function  topage(pageno){"
		response.write "var dwturl=document.URL;"
		response.write "dwturl=dwturl.replace(/page\=(\d*)/g,""page=""+pageno);"
		response.write "window.location.href=dwturl;"
		response.write "}"
		response.write "</script>"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;第<input type=""text"" size=""2"" id=""dwtpagenum""/>页"
		showpagehtml=showpagehtml & "<input type=""button"" value=""GO"" onclick=""topage(document.getElementById('dwtpagenum').value)"" />&nbsp;&nbsp;"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;"& dwtpage & "/" & dwtpagecount & "页&nbsp;&nbsp;"
		showpagehtml=showpagehtml & dwtpagesize & "条/页&nbsp;&nbsp;共" & dwtrecordcount & "条"
		
		
		
		case 2 ' 分页样式2
		showpagehtml="<table cellpadding=""8"" cellspacing=""5"" border=""1""><tr>"
		for pno=1 to dwtpagecount
			if cint(pno)=cint(dwtpage) then
				showpagehtml=showpagehtml+"<td bgcolor=""#ececec"">&nbsp;" & pno & "&nbsp;</td>"
			else
				showpagehtml=showpagehtml+"<td><a href=""" & topage(pno) & """>&nbsp;" & pno & "&nbsp;</a></td>"
			end if
		next
		showpagehtml=showpagehtml+"<td>"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;"& dwtpage & "/" & dwtpagecount & "页&nbsp;&nbsp;"
		showpagehtml=showpagehtml & dwtpagesize & "条/页&nbsp;&nbsp;共" & dwtrecordcount & "条"
		showpagehtml=showpagehtml & "</td>"
		showpagehtml=showpagehtml+"</tr></table>"
		

		
		case else
		showpagehtml="错误的dwtpagestyle参数，取值为1或2"
	
	end select
	

	response.Write showpagehtml

end sub

Private Sub Class_Terminate  '类的结束

	dwtrs.close
	set dwtrs=nothing
	
	dwtconn.close
	set dwtconn=nothing
	
end sub

end class
%>