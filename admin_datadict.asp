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


<%dim url,sqlbody,rsbody,rslevel,sqllevel,record,pgsz,total,page,rowCount,start,ii,xh
dim title


action=request("action")
url="admin_datadict.asp"


dwt.pagetop "数据字典"

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
   	'输出表单头
	'url:action的地址,forname：表单名称,title:表单标题 ,checkname:检查表单内容的名称
   dwt.lable_title "admin_datadict.asp","form1","新增数据字典内容",""   '新增
  
	'输出INPUT,
	'leftname:input在页面上显示的名称,	inputname:input在表单中的名称,inputformvalue:input的值在表单中传递用(isdisabled为真时才用，否则为空),inputvaluename:input的值在页面中显示,isdisabled:input是否为禁用,isbt:是否必添项,tips:提示信息
   dwt.lable_input "标题","title","","",false,false,""
   dwt.lable_input "描述信息","info","","",false,false,""
   dwt.lable_input "索引号","numb","","",false,false,""
   dwt.lable_input "备注","bz","","",false,false,""
 	
	'输出表单尾
	'action:action的名称,submitname:按钮的名称,isid:是否带有ID参数用于编辑修改,idname:还的ID的NAME(isid为true时添写),ID:标识ID(isid为true时添写)
   dwt.lable_footer "saveadd","添加",false,"",""
end sub	

sub saveadd()    
	 dim rsadd,sqladd
	  '保存
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
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>内容页分类管理</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
    '用户管理首页
	Dwt.out "<Div class='x-toolbar'><Div align=left><a href='admin_datadict.asp?action=add'>添加内容</a></Div></Div>" & vbCrLf
 		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf

	  Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
      Dwt.out "<tr  class=""x-grid-header"">" 
      Dwt.out "     <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""5%""><Div align=""center""><strong>ID号</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>标题</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>描述信息</strong></Div></td>"
      Dwt.out "      <td  class='x-td' style=""border-bottom-style: solid;border-width:1px"" width=""15%""><Div align=""center""><strong>索引号</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>备注</strong></Div></td>"
     Dwt.out "      <td  class='x-td' width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>选项</strong></Div></td>"
      Dwt.out "    </tr>"
      sqlbody="SELECT * from datadict ORDER BY title"
      set rsbody=server.createobject("adodb.recordset")
      rsbody.open sqlbody,connl,1,1
      if rsbody.eof and rsbody.bof then 
           Dwt.out "<p align=""center"">暂无内容</p>" 
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
                  Dwt.out "      <td width=""6%"" style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><a href='admin_datadict.asp?action=edit&ID="&rsbody("id")&"'>编辑</a>&nbsp;"
				 Dwt.out "  <a href='admin_datadict.asp?action=del&ID="&rsbody("id")&"' onClick=""return confirm('确定要删除吗？删除后对应内容将无法显示');"">删除</a></Div></td>"
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
     '编辑
	 dim id,rsedit,sqledit
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from datadict where id="&id
   rsedit.open sqledit,connl,1,1
   	'输出表单头
	'url:action的地址,forname：表单名称,title:表单标题 ,checkname:检查表单内容的名称
   dwt.lable_title "admin_datadict.asp","form1","编辑数据字典内容",""   '新增
  
	'输出INPUT,
	'leftname:input在页面上显示的名称,	inputname:input在表单中的名称,inputformvalue:input的值在表单中传递用(isdisabled为真时才用，否则为空),inputvaluename:input的值在页面中显示,isdisabled:input是否为禁用,isbt:是否必添项,tips:提示信息
   dwt.lable_input "标题","title","",rsedit("title"),false,false,""
   dwt.lable_input "描述信息","info","",rsedit("info"),false,false,""
   dwt.lable_input "索引号","numb","",rsedit("numb"),false,false,""
   dwt.lable_input "备注","bz","",rsedit("bz"),false,false,""
 	
	'输出表单尾
	'action:action的名称,submitname:按钮的名称,isid:是否带有ID参数用于编辑修改,idname:还的ID的NAME(isid为true时添写),ID:标识ID(isid为true时添写)
   dwt.lable_footer "saveedit","保存",true,"id",id

    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
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