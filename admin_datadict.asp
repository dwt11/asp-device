<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh
dim title


action=request("action")
url="admin_datadict.asp"


dwt.pagetop "�����ֵ�"

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
   	'�����ͷ
	'url:action�ĵ�ַ,forname��������,title:������ ,checkname:�������ݵ�����
   dwt.lable_title "admin_datadict.asp","form1","���������ֵ�����",""   '����
  
	'���INPUT,
	'leftname:input��ҳ������ʾ������,	inputname:input�ڱ��е�����,inputformvalue:input��ֵ�ڱ��д�����(isdisabledΪ��ʱ���ã�����Ϊ��),inputvaluename:input��ֵ��ҳ������ʾ,isdisabled:input�Ƿ�Ϊ����,isbt:�Ƿ������,tips:��ʾ��Ϣ
   dwt.lable_input "����","title","","",false,false,""
   dwt.lable_input "������Ϣ","info","","",false,false,""
   dwt.lable_input "������","numb","","",false,false,""
   dwt.lable_input "��ע","bz","","",false,false,""
 	
	'�����β
	'action:action������,submitname:��ť������,isid:�Ƿ����ID�������ڱ༭�޸�,idname:����ID��NAME(isidΪtrueʱ��д),ID:��ʶID(isidΪtrueʱ��д)
   dwt.lable_footer "saveadd","���",false,"",""
end sub	

sub saveadd()    
	 dim rsadd,sqladd
	  '����
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from datadict" 
      rsadd.open sqladd,connl,1,3
      rsadd.addnew
      rsadd("title")=ReplaceBadChar(Trim(Request("title")))
      rsadd("info")=ReplaceBadChar(Trim(Request("info")))
      rsadd("numb")=ReplaceBadChar(Trim(Request("numb")))
      rsadd("bz")=ReplaceBadChar(Trim(Request("bz")))	  
	  rsadd.update
      rsadd.close
      set rsadd=nothing
	  Dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub

sub main()
     	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>����ҳ�������</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
    '�û�������ҳ
	Dwt.out "<Div class='x-toolbar'><Div align=left><a href='admin_datadict.asp?action=add'>�������</a></Div></Div>" & vbCrLf
 		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf

	  Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      Dwt.out "<tr  class=""x-grid-header"">" 
      Dwt.out "     <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center""><strong>ID��</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>����</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>������Ϣ</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>������</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>��ע</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>ѡ��</strong></Div></td>"
      Dwt.out "    </tr>"
      sqlbody="SELECT * from datadict ORDER BY title"
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,connl,1,1
      if rsbody.eof and rsbody.bof then 
           Dwt.out "<p align=""center"">��������</p>" 
      else
           record=rsbody.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsbody.PageSize = Cint(PgSz) 
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
           rsbody.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsbody.PageSize
           do while not rsbody.eof and rowcount>0
              
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center"">"&rsbody("id")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("title")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("info")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("numb")&"</Div></td>"
                 Dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center"">"&rsbody("bz")&"</Div></td>"				 
                  Dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><a href='admin_datadict.asp?action=edit&ID="&rsbody("id")&"'>�༭</a>&nbsp;"
				 Dwt.out "  <a href='admin_datadict.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('ȷ��Ҫɾ����ɾ�����Ӧ���ݽ��޷���ʾ');"">ɾ��</a></Div></td>"
                 Dwt.out "    </tr>"
                 RowCount=RowCount-1
          rsbody.movenext
          loop
		Dwt.out "</table>"& vbCrLf
		call showpage1(page,url,total,record,PgSz)
		Dwt.out "</Div>"& vbCrLf
       end if
 	Dwt.out "</Div>"  
      rsbody.close
       set rsbody=nothing
       
end sub

sub edit()
     '�༭
	 dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from datadict where id="&id
   rsedit.open sqledit,connl,1,1
   	'�����ͷ
	'url:action�ĵ�ַ,forname��������,title:������ ,checkname:�������ݵ�����
   dwt.lable_title "admin_datadict.asp","form1","�༭�����ֵ�����",""   '����
  
	'���INPUT,
	'leftname:input��ҳ������ʾ������,	inputname:input�ڱ��е�����,inputformvalue:input��ֵ�ڱ��д�����(isdisabledΪ��ʱ���ã�����Ϊ��),inputvaluename:input��ֵ��ҳ������ʾ,isdisabled:input�Ƿ�Ϊ����,isbt:�Ƿ������,tips:��ʾ��Ϣ
   dwt.lable_input "����","title","",rsedit("title"),false,false,""
   dwt.lable_input "������Ϣ","info","",rsedit("info"),false,false,""
   dwt.lable_input "������","numb","",rsedit("numb"),false,false,""
   dwt.lable_input "��ע","bz","",rsedit("bz"),false,false,""
 	
	'�����β
	'action:action������,submitname:��ť������,isid:�Ƿ����ID�������ڱ༭�޸�,idname:����ID��NAME(isidΪtrueʱ��д),ID:��ʶID(isidΪtrueʱ��д)
   dwt.lable_footer "saveedit","����",true,"id",id

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'�༭����
dim rsedit,sqledit
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from datadict where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,connl,1,3
rsedit("title")=ReplaceBadChar(Trim(Request("title")))
rsedit("info")=ReplaceBadChar(Trim(Request("info")))
rsedit("numb")=ReplaceBadChar(Trim(Request("numb")))
rsedit("bz")=ReplaceBadChar(Trim(Request("bz")))	  
rsedit.update
rsedit.close
	
	  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
dim id,rsdel,sqldel
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from datadict where id="&id
rsdel.open sqldel,connl,1,3
Dwt.out"<Script Language=Javascript>history.go(-1)</Script>"
'rsdel.close
set rsdel=nothing  

end sub

Dwt.out "</body></html>"

Call CloseConn
%>