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



<%
dim sqlpxst,rspxst,title,record,pgsz,total,page,start,rowcount,url,ii
dim rsadd,sqladd,id,rsedit,sqledit,scontent,rsdel,sqldel
url="pxst.asp"
dim keys
keys=request("keyword")

dwt.out "<html>"& vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<title>信息管理系统试题库管理页</title>"& vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "</head>"& vbCrLf
dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

action=request("action")

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
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>添 加 试 题</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='pxst.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >标题:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=pxst_title>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >分类:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out "<select name='pxst_class'>"
	dwt.out "<option value='0'>请选择分类</option>"& vbCrLf

	dim sql,rs
	sql="SELECT * from pxst_class"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connpxjhzj,1,1
    if rs.eof then 
	else
	do while not rs.eof
       	response.write"<option value='"&rs("id")&"'>"&rs("class_name")&"</option>"& vbCrLf
	    'usernameh=rsbz("username1")
		rs.movenext
	loop
	end if 
	rs.close
	set rs=nothing
	dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&session("username1")&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&date()&"'  disabled='disabled' >"& vbCrLf
	dwt.out "<input name='pxst_date' type='hidden' value='"&date()&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	 dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=pxst_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='600' HEIGHT='350'>"
       
      dwt.out"</iframe>  <input type='hidden' name='pxst_body' value=''>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  
	
end sub	

sub saveadd()    
	  '保存
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from pxst" 
      rsadd.open sqladd,conne,1,3
      rsadd.addnew
      rsadd("pxst_title")=ReplaceBadChar(Trim(Request("pxst_title")))
      'rsadd("pxst_zz")=request("pxst_zz")
      rsadd("pxst_body")=Trim(request("pxst_body"))
      rsadd("pxst_date")=request("pxst_date")
      rsadd("pxst_class")=request("pxst_class")
      rsadd("userid")=session("userid")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out"<Script Language=Javascript>location.href='pxst.asp';</Script>"
end sub

sub edit()
     '编辑
	 
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from pxst where id="&id
   rsedit.open sqledit,conne,1,1
	dwt.out"<DIV style='WIDTH: 800px;padding-top:20px;padding-left:10px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'><div align=center>编 辑 试 题</div></H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
	dwt.out "<form method='post' class='x-form' action='pxst.asp' name='form1' >"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
				  
	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >标题:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=pxst_title value='"&rsedit("pxst_title")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px' >分类:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element >"& vbCrLf
	dwt.out "<select name='pxst_class'>"
	dwt.out "<option value='0'>请选择分类</option>"& vbCrLf

	dim sql,rs
	sql="SELECT * from pxst_class"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connpxjhzj,1,1
    if rs.eof then 
	else
	do while not rs.eof
       	response.write"<option value='"&rs("id")&"' "
		if rsedit("pxst_class")=rs("id") then dwt.out "selected"
		dwt.out ">"&rs("class_name")&"</option>"& vbCrLf
	    'usernameh=rsbz("username1")
		rs.movenext
	loop
	end if 
	rs.close
	set rs=nothing
	dwt.out"</select>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>发布人:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&usernameh(rsedit("userid"))&"'  disabled='disabled'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-clear-left'></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>时间:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' value='"&rsedit("pxst_date")&"'  disabled='disabled' >"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  
	dwt.out"							<DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 75px'>内容:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' >"& vbCrLf
	scontent=rsedit("pxst_body")
	 dwt.out"<iframe ID='eWebEditor1' src='/eweb/ewebeditor.asp?id=pxst_body&style=s_blue &originalfilename=d_originalfilename &savefilename=d_savefilename &savepathfilename=d_savepathfilename' frameborder='0' scrolling='no' width='100%' HEIGHT='350'>"
       
      dwt.out"</iframe><textarea name='pxst_body' style='display:none'>"&Server.HtmlEncode(sContent)&"</textarea>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	

	dwt.out"			  <DIV class=x-form-clear></DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
	dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"		  </FORM>"& vbCrLf
	dwt.out"		</DIV>"& vbCrLf
	dwt.out"	  </DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-bl>"& vbCrLf
	dwt.out"	<DIV class=x-box-br>"& vbCrLf
	dwt.out"	  <DIV class=x-box-bc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  


    rsedit.close
    set rsedit=nothing
end sub

sub saveedit()
'编辑保存
set rsedit=server.createobject("adodb.recordset")
sqledit="select * from pxst where ID="&ReplaceBadChar(Trim(request("ID")))

rsedit.open sqledit,conne,1,3
rsedit("pxst_title")=ReplaceBadChar(Trim(Request("pxst_title")))
rsedit("pxst_body")=Trim(request("pxst_body"))
rsedit("pxst_class")=Trim(request("pxst_class"))
rsedit.update
rsedit.close
	  dwt.out"<Script Language=Javascript>history.go(-2);</Script>"
end sub


sub main()
	dwt.out "<div style='left:6px;'>"& vbCrLf
	dwt.out "     <DIV class='x-layout-panel-hd'>"& vbCrLf
	dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>培训试题"&title&"</span>"& vbCrLf
	dwt.out "     </div>"& vbCrLf
	search	()

	sqlpxst="SELECT * from pxst" 
	if request("classid")<>"" then sqlpxst=sqlpxst&" where pxst_class="&request("classid")
	if keys<>"" then sqlpxst=sqlpxst&" where pxst_body like '%" &keys& "%' "
	sqlpxst=sqlpxst&" ORDER BY id DESC"
	set rspxst=server.createobject("adodb.recordset")
	rspxst.open sqlpxst,conne,1,1
	if rspxst.eof and rspxst.bof then 
	message("未找到相关培训试题")
	else
		dwt.out "<DIV class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		dwt.out "<tr class=""x-grid-header"">" 
		dwt.out "     <td  class='x-td' ><DIV class='x-grid-hd-text'>序号</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>分    类</div></td>"
		dwt.out "      <td class='x-td' ><DIV class='x-grid-hd-text'>标    题</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布者</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>发布时间</div></td>"
		dwt.out "      <td  class='x-td'><DIV class='x-grid-hd-text'>选项</div></td>"
		dwt.out "    </tr>"
           record=rspxst.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxst.PageSize = Cint(PgSz) 
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
           rspxst.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxst.PageSize
           do while not rspxst.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" >"&class_name(rspxst("pxst_class"))&"</td>"
                 dwt.out "      <td style=""border-bottom-style: solid;border-width:1px"" ><a href=pxst_view.asp?id="&rspxst("id")&" target=_blank>"&searchH(uCase(rspxst("pxst_title")),keys)&"</a></td>"
                 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 if isnull(rspxst("userid")) then 
				   dwt.out rspxst("pxst_zz")
				 else
				   dwt.out usernameh(rspxst("userid")) 
				 end if   
				 dwt.out"&nbsp;</div></td>"
                 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_date")&"</div></td>"
                 dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"
				 if session("level")=0 or rspxst("userid")=session("userid") then
				  dwt.out "<a href='pxst.asp?action=edit&ID="&rspxst("id")&"'>编辑</a>"
				  dwt.out "&nbsp;<a href='pxst.asp?action=del&ID="&rspxst("id")&"' onClick=""return confirm('确定要删除此试题吗？');"">删除</a>"
				 end if 			'call editdel(rspxst("id"),rspxst("sscj"),"pxst.asp?action=edit&id=","pxst.asp?action=del&id=")
				 dwt.out "&nbsp; </div></td>"
                 dwt.out "    </tr>"
                 RowCount=RowCount-1
          rspxst.movenext
          loop
		dwt.out "</table>"& vbCrLf
		if keys<>"" then
		  call showpage(page,url,total,record,PgSz)
		else
		  call showpage1(page,url,total,record,PgSz)
		end if 
		dwt.out "</div>"& vbCrLf
	end if
	dwt.out "</div>"  
	rspxst.close
	set rspxst=nothing
	conn.close
	set conn=nothing
end sub



sub del()
ID=request("ID")
set rsdel=server.createobject("adodb.recordset")
sqldel="delete * from pxst where id="&id
rsdel.open sqldel,conne,1,3
dwt.out"<Script Language=Javascript>history.go(-1);</Script>"
'rsdel.close
set rsdel=nothing  

end sub


Function class_name(class_id)
    dim sqlcj,rscj
'dim class_id

	  sqlcj="SELECT * from pxst_class where id="&class_id
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,connpxjhzj,1,1
    if rscj.eof then 
		class_name="未编辑"
	else
	do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    class_name=rscj("class_name")
		rscj.movenext
	loop
	end if 
	rscj.close
	set rscj=nothing
end Function
dwt.out "</body></html>"


sub search()
dim sqlcj,rscj
dwt.out "<div class='x-toolbar'>" & vbCrLf
dwt.out "<form method='Get' name='SearchForm' action='pxst.asp'>" & vbCrLf
dwt.out "<a href=""pxst.asp?action=add"">添加试题</a>&nbsp;&nbsp;标题搜索：" & vbCrLf
dwt.out "  <input type='text' name='keyword' size='20' maxlength='50' value="&keys&">" & vbCrLf
dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf
dwt.out "</form></div>" & vbCrLf
end sub

Call CloseConn
%>