<%
class dwt_page
'��ҳ��

public  dwtconn    	 'connection���ݿ����Ӷ���
public  dwtrs      	 'recordset����
public  dwttable   	 'Ҫ��ҳ��table���ݿ��
public  dwtpagesize      '��ҳ��pagesize��С
public  dwtpage          'URL�����page����ֵ
public  dwtrssql         '��ҳ��SQL���
public  dwttourl         '���ӵ�URL
public  dwtrecordcount   '���ݿ��¼��
public  dwtpagecount     '��ҳ����
public  dwtpagestyle     '��ҳ����ʽ��ȡֵΪ1��2����ʽ1�ʺ�ҳ���϶�ģ���ʽ2�ʺ�ҳ�����ٵ�
public  dwtsql    '���ݼ���ʱ������,����ʱ������" where ���ʽ and "ǰ��ո�Ҫ��

Private Sub Class_Initialize '��ĳ�ʼ��

	dwtpagesize=10 'Ĭ�Ϸ�ҳ��С
	dwtpagestyle=1 'Ĭ�Ϸ�ҳ��ʾ��ʽ
	
	if isempty(request.querystring("page")) <> true then '��֤page����Ĳ���
	
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

Public Property Let getconn(zwconn) '�õ����ݿ����ӵ�connection����
	
	set dwtconn=zwconn
	
End Property 

public sub dwt_set '�����ṩ�Ĳ�����������

	dwtrecordcount=dwtconn.execute("select count(id) from " & dwttable&replace(dwtsql,"and",""))(0) '���ݿ��ܼ�¼��
	
	if dwtrecordcount mod dwtpagesize=0 then '������ҳ��
	
		dwtpagecount=int(dwtrecordcount/dwtpagesize)
		
	else
	
		dwtpagecount=int(dwtrecordcount/dwtpagesize)+1
		
	end if
	
	if clng(dwtpage) > clng(dwtpagecount) then '���page�����ҳ�����������ҳ�룬ȡ���ҳ��
	
		dwtpage=dwtpagecount
		
	end if
	
	dwtrssql="select top " & dwtpagesize  & " * from  " & dwttable '��̬�齨��ҳ��SQL���
	dwtrssql=dwtrssql & " where "&replace(dwtsql,"where","")&" ID<=( select min(ID) from (select top " & cint(dwtpagesize)*cint(dwtpage-1)+1
	dwtrssql=dwtrssql & " ID from " & dwttable & " order by id desc) as tabview) order by  id desc"
	
	set dwtrs=dwtconn.execute(dwtrssql) '����һ��ֻ����ǰ��recordset����
	
end sub


private function topage(n) '��ҳ

	dwttourl="?page=" & n 
	
	for each str in request.querystring 'URL��������в���
	
			if lcase(str) <> "page" then
			
				dwttourl=dwttourl & "&" & str & "=" & request.querystring(str)
				
			end if
	
	next
	
	topage=dwttourl
	
end function

public sub dwtshowpage '��ʾ��ҳ

	dim showpagehtml
	
	select case dwtpagestyle
		
		case 1 '��ҳ��ʽ1
		if cint(dwtpage)<=1 then
		
			showpagehtml="��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;<a href=""" & topage(dwtpage+1) & """>"
			showpagehtml=showpagehtml & "��һҳ</a>&nbsp;&nbsp;<a href=""" & topage(dwtpagecount) & """>βҳ</a>"
			
		elseif cint(dwtpage)>=dwtpagecount then
		
			showpagehtml="<a href=""" & topage(1) & """>��ҳ</a>&nbsp;&nbsp;<a href=""" & topage(dwtpage-1) & """>��һҳ</a>"
			showpagehtml=showpagehtml & "&nbsp;&nbsp;��һҳ&nbsp;&nbsp;βҳ"
			
		else
		
			showpagehtml="<a href=""" & topage(1) & """>��ҳ</a>&nbsp;&nbsp;<a href=""" & topage(dwtpage-1) & """>��һҳ</a>&nbsp;&nbsp;"
			showpagehtml=showpagehtml & "<a href=""" & topage(dwtpage+1) & """>��һҳ</a>&nbsp;&nbsp;<a href=""" & topage(dwtpagecount) & """>βҳ</a>"
			
		end if
		
		response.write "<script>"
		response.write "function  topage(pageno){"
		response.write "var dwturl=document.URL;"
		response.write "dwturl=dwturl.replace(/page\=(\d*)/g,""page=""+pageno);"
		response.write "window.location.href=dwturl;"
		response.write "}"
		response.write "</script>"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;��<input type=""text"" size=""2"" id=""dwtpagenum""/>ҳ"
		showpagehtml=showpagehtml & "<input type=""button"" value=""GO"" onclick=""topage(document.getElementById('dwtpagenum').value)"" />&nbsp;&nbsp;"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;"& dwtpage & "/" & dwtpagecount & "ҳ&nbsp;&nbsp;"
		showpagehtml=showpagehtml & dwtpagesize & "��/ҳ&nbsp;&nbsp;��" & dwtrecordcount & "��"
		
		
		
		case 2 ' ��ҳ��ʽ2
		showpagehtml="<table cellpadding=""8"" cellspacing=""5"" border=""1""><tr>"
		for pno=1 to dwtpagecount
			if cint(pno)=cint(dwtpage) then
				showpagehtml=showpagehtml+"<td bgcolor=""#ececec"">&nbsp;" & pno & "&nbsp;</td>"
			else
				showpagehtml=showpagehtml+"<td><a href=""" & topage(pno) & """>&nbsp;" & pno & "&nbsp;</a></td>"
			end if
		next
		showpagehtml=showpagehtml+"<td>"
		showpagehtml=showpagehtml & "&nbsp;&nbsp;"& dwtpage & "/" & dwtpagecount & "ҳ&nbsp;&nbsp;"
		showpagehtml=showpagehtml & dwtpagesize & "��/ҳ&nbsp;&nbsp;��" & dwtrecordcount & "��"
		showpagehtml=showpagehtml & "</td>"
		showpagehtml=showpagehtml+"</tr></table>"
		

		
		case else
		showpagehtml="�����dwtpagestyle������ȡֵΪ1��2"
	
	end select
	

	response.Write showpagehtml

end sub

Private Sub Class_Terminate  '��Ľ���

	dwtrs.close
	set dwtrs=nothing
	
	dwtconn.close
	set dwtconn=nothing
	
end sub

end class
%>