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
'�ļ���β��Ȩ
'***********************888
sub footer()
response.write "<br><br>"
response.write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""100%"" class=""border"" align=center>"
response.write "<tr align=""center"">"
response.write "<td height=25 class=""topbg""><span class=""Glow"">�豸����ϵͳ All Rights Reserved.</span>"
response.write "</tr></table></body></html>"
end sub



'**************************************************
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
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
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
'URL�д�����
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
   response.write "</div></td></tr></table>"
end sub


'�ٷ�֮80�����ʾ
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
   response.write "</div></td></tr></table>"
end sub


'********************************************8
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
 'url�в�����
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
   response.write"&nbsp;&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&total&"</strong>ҳ&nbsp;&nbsp;"
   if Total =1 then 
     response.write"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3'  disabled='disabled' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">��"
   else
     response.write"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">��"
   end if 
   if Total=1 then 
       response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
       response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >��"&page&"ҳ</option>"
     else
    	 response.write"  <option value='"&ii&"'>��"&ii&"ҳ</option>"
     end if 
   next 
   response.write" </select>&nbsp;&nbsp;��"&record&"������"
   response.write "</div></td></tr></table>"
end sub



'1ά��һ���䣬2ά�޶����䣬3ά�������䣬4ά���ĳ��䣬5�ۺϳ��䣬6������
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

'���ڶ̵ĳ�����ʾ
Function sscjh_d(sscj)
       dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    sscjh_d=replace(replace(replace(rscj("levelname"),"��",""),"����",""),"��","")
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
end Function 

 '29���¼�
'���ڱ༭����װ����ʾ
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
		gh="δ�༭"
	else
		response.write"<option value='0'"
		if ghid=0 then response.write " selected" 
		response.write">��ѡ��װ��</option>"
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
 '29���¼�
'ȡװ�ù�������
Function gh(ghid)
       dim sqlgh,rsgh
	if isnull(ghid) then ghid=0
	sqlgh="SELECT * from ghname where ghid="&ghid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    if rsgh.eof then 
	  gh="δ�༭"
	else
	    gh=rsgh("gh_name")
end if 
	rsgh.close
	set rsgh=nothing
end Function 
 '29���¼�
'ȡ�ּ�������
Function fj(fjnumb)
       dim fj_i
	if isnull(fjnumb) or fjnumb=0 then 
	  fj="δ�ּ�"
	else
		for fj_i=1 to fjnumb
		fj=fj&"*"
		next
	  
	end if 
end Function 





'�ȶ���1,��ˮ��2,�ϳ�һ��3,�ϳɶ���4,��ѹ��5,���ʰ�6,��藺�7,�����8
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


'ѡ��༭��ɾ����
sub editdel(id,sscj,editurl,delurl)
 if session("level")=sscj or session("level")=0 then 
    response.write "<a href="&editurl&id&">�༭</a>&nbsp;"
	response.write "<a href="&delurl&id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ��</a>"
 else
    response.write "&nbsp;"
 end if 
end sub

'ѡ��ࡢɾ��
sub editdel_d(id,sscj,editurl,delurl)
 if session("level")=sscj or session("level")=0 then 
    response.write "<a href="&editurl&id&">��</a>&nbsp;"
	response.write "<a href="&delurl&id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
 else
    response.write "&nbsp;"
 end if 
end sub
 
 
 '������ʾ�����ؼ���
function searchH(body,key)
dim inum'��һ�γ��ֵ�λ��
dim leftbody '��ȡBODY�е��ַ�����INUM
dim lenkey   'KEY���ַ�������
dim lenbody 'BODY���ַ�������
dim rightbody '��ȡBODY�е��ַ�����KEY��β��
dim midikey  '��KEY������ʾΪ��ɫ
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

'�����ӱ�����ʾѡ�񳵼�ı�
function formsscj()
   if session("level")=0 then 
	response.write"<select name='sscj' size='1'>"
    response.write"<option >��ѡ����������</option>"
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


'�����ӱ�����ʾѡ�񳵼�Ͱ���ı�
function formsscjbz()
 dim rscj,sqlcj,rsbz,sqlbz
 if session("level")=0 then 
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	response.write"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    response.write"<option  selected>ѡ����������</option>"& vbCrLf
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
    response.write "<option  selected>ѡ��������</option>" & vbCrLf
    response.write "</select></td></tr>  "  & vbCrLf
    response.write "<script><!--" & vbCrLf
    response.write "var groups=document.form1.sscj.options.length" & vbCrLf
    response.write "var group=new Array(groups)" & vbCrLf
    response.write "for (i=0; i<groups; i++)" & vbCrLf
    response.write "group[i]=new Array()" & vbCrLf
    response.write "group[0][0]=new Option(""ѡ��������"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0	
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   response.write "group["&rscj("levelid")&"][0]=new Option(""δ��Ӱ���"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'response.write"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
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
   	  response.write"<option value='0'>δ��Ӱ���</option>"
   else   
	  'response.write"<option value='0'>����</option>"
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


'ȡ��ǰ��ҳURL
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
	'�ж�URL���Ƿ��з�ҳ����������ȥ��
	if InStr(strtemp,"pgsz")<>0 then
		urllen=InStr(strtemp,"pgsz")
		strtemp=left(strtemp,urllen-2)
	end if  
	GetUrl = strtemp 
End Function


function message(text)
response.write "<br><br><br><table width=""351"" height=""185"" border='0' align='center' cellpadding='2' cellspacing='1' class='border' >" & vbCrLf
response.write "  <tr  class='title'>" & vbCrLf
response.write "    <td height=""30"" colspan=""3""><div align=center>ϵͳ��ʾ��Ϣ</div></td>" & vbCrLf
response.write "  </tr>" & vbCrLf
response.write "  <tr>" & vbCrLf
response.write "    <td><div align=""center""><img src=""/images/3.gif"" width=""60"" height=""60""></div></td>" & vbCrLf
response.write "    <td>&nbsp;</td>" & vbCrLf
response.write "    <td>"&text&"</td>" & vbCrLf
response.write "  </tr>" & vbCrLf
response.write "</table>" & vbCrLf

end function
%>