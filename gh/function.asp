    
    <%
	'��ȡ��������
	Function getclassname(classid)
       dim sqlgh,rsgh
	if isnull(classid) then classid=0
	sqlgh="SELECT * from dgtzl_index_gh where id="&classid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conndgt,1,1
    if rsgh.eof then 
	  getclassname="δ�༭"
	else
	    'getclassname="<a href=showlist.asp?classid="&classid&">"&rsgh("class_name")&"</a>"
		getclassname=rsgh("class_name")
end if 
	rsgh.close
	set rsgh=nothing
end Function 


	'��ȡ����
	Function gettitlename(bodyid)
       dim sqlgh,rsgh
	if isnull(bodyid) then bodyid=0
	sqlgh="SELECT * from dgtzl_body where id="&bodyid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conndgt,1,1
    if rsgh.eof then 
	  gettitlename="δ�༭"
	else
	    gettitlename=rsgh("news_title")
	end if 
		rsgh.close
		set rsgh=nothing
end Function 






	'ѭ����ȡ�ϼ����������ڵ�ǰ��ʾ
	Function gettclassname(classid)
       dim sqlgh,rsgh
	if classid="" then 
	else
		sqlgh="SELECT * from dgtzl_index_gh where id="&classid
		set rsgh=server.createobject("adodb.recordset")
		rsgh.open sqlgh,conndgt,1,1
		if rsgh.eof then 
		    gettclassname="> δ�༭"
		else
		
		  if rsgh("index")<>0 then 
			  sqlgh1="SELECT * from dgtzl_index_gh where id="&rsgh("index")
			  set rsgh1=server.createobject("adodb.recordset")
			  rsgh1.open sqlgh1,conndgt,1,1
			  if rsgh1.eof then 
				  gettclassname="> δ�༭"
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
	
	
	
	'��ȡ������ ��-�� 
	Function getsdate(newdate)
	getsdate=month(newdate)&"-"&day(newdate)
	end Function 
	
	'��ȡ��-��-�� 
	Function gettdate(newdate)
	gettdate=year(newdate)&"-"&month(newdate)&"-"&day(newdate)
	end Function 
	
	'��ʽ ��������ͼƬ��С 
	  function imgCode(strContent)
	  dim re
	  Set re=new RegExp
	  re.IgnoreCase =true
	  re.Global=True
			  
	  re.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
	  strContent=re.replace(strContent,"<div align=center><img SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"" border=0 alt=�������´������ͼƬ onload=""javascript:if(this.width>550)this.width=333""></div>")
	  
	  
	  set re=Nothing
	  imgCode=strContent
	  end function	
	
	'********************************************8
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
'URL�д�����
'*******************************************
sub newshowpage(page,url,total,record,pgsz)
   response.write "<div class='page'>"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1>��ҳ</a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&">��һҳ</a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&">��һҳ</a> <a href="&url&"&pgsz="&pgsz&"&page="&total&">βҳ</a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&total&"</strong>ҳ&nbsp;&nbsp;"
  if Total =1 then 
    response.write"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">��"
  else
   response.write"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">��"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >��"&page&"ҳ</option>"
     else
    	 response.write"  <option value='"&ii&"'>��"&ii&"ҳ</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;��"&record&"������"
   response.write "</div>"
end sub


	%>