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
dim url,lb,brxx,sqlfhqc,rsfhqc,record,pgsz,total,page,start,rowcount,ii
dim rsadd,sqladd,id,rsdel,sqldel,rsedit,sqledit
url="fhqc.asp"

dwt.out  "<html>"& vbCrLf
dwt.out  "<head>" & vbCrLf
dwt.out  "<title>信息管理系统防护器材管理页</title>"& vbCrLf
dwt.out  "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
dwt.out  "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out  "<SCRIPT language=javascript>" & vbCrLf
dwt.out  "function checkadd(){" & vbCrLf
dwt.out  "if(document.form1.fhqc_sscj.value==''){" & vbCrLf
dwt.out  "alert('请选择所属车间！');" & vbCrLf
dwt.out  "document.form1.fhqc_sscj.focus();" & vbCrLf
dwt.out  "return false;" & vbCrLf
dwt.out  "}" & vbCrLf
dwt.out  "}" & vbCrLf
dwt.out  "</SCRIPT>" & vbCrLf
dwt.out  "</head>"& vbCrLf
dwt.out  "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
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
dim rscj,sqlcj,rs,sql
   dwt.out "<br><br><br><form method='post' action='fhqc.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>添加防护器材</strong></div></td>    </tr>"
   dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间班组： </strong></td>"      
   dwt.out "<td width='88%' class='tdbg'>"
  if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	dwt.out "<select name='fhqc_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    dwt.out "<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	dwt.out "<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    dwt.out "</select>"  	 & vbCrLf
    dwt.out  "<select name='fhqc_ssbz' size='1' >" & vbCrLf
    dwt.out  "<option  selected>选择班组分类</option>" & vbCrLf
    dwt.out  "</select></td></tr>  "  & vbCrLf
    dwt.out  "<script><!--" & vbCrLf
    dwt.out  "var groups=document.form1.fhqc_sscj.options.length" & vbCrLf
    dwt.out  "var group=new Array(groups)" & vbCrLf
    dwt.out  "for (i=0; i<groups; i++)" & vbCrLf
    dwt.out  "group[i]=new Array()" & vbCrLf
    dwt.out  "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=1		
	 dim sqlbz,rsbz
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   dwt.out  "group["&rscj("levelid")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   dwt.out "group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   dwt.out "group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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




    dwt.out  "var temp=document.form1.fhqc_ssbz" & vbCrLf
    dwt.out  "function redirect(x){" & vbCrLf
    dwt.out  "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    dwt.out  "temp.options[m]=null" & vbCrLf
    dwt.out  "for (i=0;i<group[x].length;i++){" & vbCrLf
    dwt.out  "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    dwt.out  "}" & vbCrLf
    dwt.out  "temp.options[0].selected=true" & vbCrLf
    dwt.out  "}//--></script>" & vbCrLf



  else 	 
   dwt.out "<input name='fhqc_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   dwt.out "<input name='fhqc_sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   sql="SELECT * from bzname where sscj="&session("levelclass")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   dwt.out "<select name='fhqc_ssbz' size='1'>"
   
   if rs.eof and rs.bof then 
   	  dwt.out "<option value='0'>车间</option>"
   else   
	  dwt.out "<option value='0'>车间</option>"
      do while not rs.eof
	     dwt.out "<option value='"&rs("id")&"'>"&rs("bzname")&"</option>"
	  rs.movenext
      loop
	end if 
	 dwt.out "</select>" 
  rs.close
  set rs=nothing
end if 
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out "<strong>器材名称：</strong></td>"
	 dwt.out "<td width='88%' class='tdbg'><input name='fhqc_name' type='text'></td>    </tr>   "
	 
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单&nbsp;&nbsp;位：</strong></td> "
	 dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_dw' ></td></tr> "
	 
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数&nbsp;&nbsp;量：</strong></td> "
	dwt.out "<td><input type='text' name='fhqc_numb' ></td></tr>"
	 
		dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>发放时间：</strong></td> "
   dwt.out "<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out "<input name='fhqc_date' type='text' value="&now()&" >"
   dwt.out "<a href='#' onClick=""popUpCalendar(this, fhqc_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out "<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>领取人：</strong></td>"      
    dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_lqrname'></td></tr>  "   


	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_bz'></td></tr>  "   

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveadd'>    <input  type='submit' name='Submit' value=' 添 加 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
end sub	

sub saveadd()    
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from fhqc" 
      rsadd.open sqladd,connb,1,3
      rsadd.addnew
on error resume next
	  rsadd("sscj")=Trim(Request("fhqc_sscj"))
      rsadd("ssbz")=Trim(Request("fhqc_ssbz"))
	  rsadd("date")=request("fhqc_date")
      rsadd("name")=Trim(request("fhqc_name"))
      rsadd("dw")=request("fhqc_dw")
      rsadd("numb")=request("fhqc_numb")
      rsadd("lqrname")=request("fhqc_lqrname")
      rsadd("bz")=request("fhqc_bz")
      rsadd.update
      rsadd.close
      set rsadd=nothing
	  dwt.out "<Script Language=Javascript>location.href='fhqc.asp';</Script>"
end sub

sub saveedit()    
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from fhqc where id="&ReplaceBadChar(Trim(request("id")))
      rsedit.open sqledit,connb,1,3
 on error resume next
     rsedit("sscj")=Trim(Request("fhqc_sscj"))
      rsedit("ssbz")=request("fhqc_ssbz")
      rsedit("date")=Trim(request("fhqc_date"))
      rsedit("name")=request("fhqc_name")
      rsedit("dw")=request("fhqc_dw")
      rsedit("numb")=request("fhqc_numb")
      rsedit("bz")=request("fhqc_bz")
	  rsedit("lqrname")=request("fhqc_lqrname")
      rsedit.update
      rsedit.close
      set rsedit=nothing
	  dwt.out "<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub del()
  id=request("id")
  set rsdel=server.createobject("adodb.recordset")
  sqldel="delete * from fhqc where id="&id
  rsdel.open sqldel,connb,1,3
  dwt.out "<Script Language=Javascript>history.go(-1)</Script>"
  set rsdel=nothing  
end sub


sub edit()
   id=ReplaceBadChar(Trim(request("id")))
   set rsedit=server.createobject("adodb.recordset")
   sqledit="select * from fhqc where id="&id
   rsedit.open sqledit,connb,1,1
   dwt.out "<br><br><br><form method='post' action='fhqc.asp' name='form1' onsubmit='javascript:return checkadd();'>"
   dwt.out "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
   dwt.out "<tr class='title'><td height='22' colspan='2'>"
   dwt.out "<div align='center'><strong>编辑防护器材表</strong></div></td>    </tr>"
     dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>所属车间： </strong></td>"   & vbCrLf   
     dwt.out "<td width='88%' class='tdbg'><input name='fhqc_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sscj"))&"'></td></tr>"& vbCrLf
     dwt.out "<input name='fhqc_sscj' type='hidden' value="&rsedit("sscj")&"></td></tr>"& vbCrLf

dwt.out "<tr class='tdbg'><td width='15%' align='right' class='tdbg'><strong>所属班组： </strong></td>"& vbCrLf      
    dwt.out "<td width='88%' class='tdbg'>"& vbCrLf
	dim ssbz
	if rsedit("ssbz")=0 then
  	   ssbz="车间"
	else
	   ssbz=ssbzh(rsedit("ssbz"))
	end if    
    dwt.out "<input name=""fhqc_ssbz"" value="&ssbz&" type='text' disabled='disabled' >"& vbCrLf
     dwt.out "<input name='fhqc_ssbz' type='hidden' value="&rsedit("ssbz")&"></td></tr>"& vbCrLf

	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 dwt.out "<strong>器材名称：</strong></td>"
	 dwt.out "<td width='88%' class='tdbg'><input name='fhqc_name' type='text' value="&rsedit("name")&"></td>    </tr>   "
	 
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>单&nbsp;&nbsp;位：</strong></td> "
	 dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_dw' value="&rsedit("dw")&" ></td></tr> "
	 
	 dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>数&nbsp;&nbsp;量：</strong></td> "
	dwt.out "<td><input type='text' name='fhqc_numb'  value="&rsedit("numb")&"></td></tr>"
	 
		dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>发放时间：</strong></td> "
   dwt.out "<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
   dwt.out "<input name='fhqc_date' type='text'  value="&rsedit("date")&">"
   dwt.out "<a href='#' onClick=""popUpCalendar(this, fhqc_date, 'yyyy-mm-dd'); return false;"">"
   dwt.out "<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a></td></tr>"& vbCrLf
	
	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>领取人：</strong></td>"      
    dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_lqrname'  value="&rsedit("lqrname")&"></td></tr>  "   


	dwt.out "<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>备&nbsp;&nbsp;注：</strong></td>"      
    dwt.out "<td width='88%' class='tdbg'><input type='text' name='fhqc_bz' value="&rsedit("bz")&"></td></tr>  "   

	dwt.out "<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	dwt.out "<input name='action' type='hidden' id='action' value='saveedit'> <input type='hidden' name='id' value='"&id&"'>      <input  type='submit' name='Submit' value=' 保  存 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	dwt.out "</table></form>"
	       rsedit.close
       set rsedit=nothing
	

end sub


sub main()
dwt.out  "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
dwt.out  "<tr class='topbg'>"& vbCrLf
dwt.out  "<td height='22' colspan='2' align='center'><strong>防护器材管理页</strong></td>"& vbCrLf
dwt.out  "</tr>  "& vbCrLf
dwt.out  "<tr class='tdbg'>"& vbCrLf
dwt.out  "<td width='70' height='30'><strong>管理导航：</strong></td>"& vbCrLf
dwt.out  "<td height='30'><a href=""fhqc.asp"">防护器材首页</a>&nbsp;|&nbsp;<a href=""fhqc.asp?action=add"">添加防护器材</a></td>"& vbCrLf
dwt.out  "</tr>"& vbCrLf
dwt.out  "</table>"& vbCrLf

sqlfhqc="SELECT * from fhqc ORDER BY id DESC"
set rsfhqc=server.createobject("adodb.recordset")
rsfhqc.open sqlfhqc,connb,1,1
if rsfhqc.eof and rsfhqc.bof then 
message("未添加相关内容")
else

dwt.out  "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">" & vbCrLf
dwt.out  "<tr class=""title"">"  & vbCrLf
dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>序号</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>领用时间</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>领用单位</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>器材名称</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>单位</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>数量</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>领取人</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>备注</strong></div></td>" & vbCrLf
dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>选项</strong></div></td>" & vbCrLf
dwt.out  "    </tr>" & vbCrLf

		   record=rsfhqc.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rsfhqc.PageSize = Cint(PgSz) 
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
           rsfhqc.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rsfhqc.PageSize
           do while not rsfhqc.eof and rowcount>0
                dwt.out  "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">" & vbCrLf
                dwt.out  "     <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("id")&"</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("date")&"</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&sscjh_d(rsfhqc("sscj"))&ssbzh(rsfhqc("ssbz"))&"</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px"">"&rsfhqc("name")&"&nbsp;</td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("dw")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("numb")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("lqrname")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rsfhqc("bz")&"&nbsp;</div></td>" & vbCrLf
                dwt.out  "      <td style=""border-bottom-style: solid;border-width:1px""><div align=center>" & vbCrLf
				call editdel(rsfhqc("id"),rsfhqc("sscj"),"fhqc.asp?action=edit&id=","fhqc.asp?action=del&id=")
                dwt.out  "</div></td></tr>" & vbCrLf
                 RowCount=RowCount-1
          rsfhqc.movenext
          loop
        dwt.out  "</table>" & vbCrLf
       call showpage1(page,url,total,record,PgSz)
	   end if
       rsfhqc.close
       set rsfhqc=nothing
        conn.close
        set conn=nothing
end sub





dwt.out  "</body></html>"


Call CloseConn
%>