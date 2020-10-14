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
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>计量管理管理页</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.Out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.Out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function check(){" & vbCrLf

Dwt.out "if(document.form1.sscj.value==''){" & vbCrLf
Dwt.out "alert('请选择所属车间！');" & vbCrLf
Dwt.out "document.form1.sscj.focus();" & vbCrLf
Dwt.out "return false;" & vbCrLf
Dwt.out "}" & vbCrLf

Dwt.out "}" & vbCrLf

Dwt.out "function complete(){" & vbCrLf

Dwt.out "if(document.form2.zjinfo.value==''){" & vbCrLf
Dwt.out "alert('周检结果未添写！');" & vbCrLf
Dwt.out "document.form2.zjinfo.focus();" & vbCrLf
Dwt.out "return false;" & vbCrLf
Dwt.out "}" & vbCrLf

Dwt.out "}" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.Out"<script language=javascript src='/js/popselectdate.js'></script>"

Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
dim sqlcj,rscj,i,ii,sqlbz,rsbz,sql,rs,sqld,rsd
    dim url,record,pgsz,total,page,start,rowcount
	dim zjyear,zjmonth
	dim sscj,ssbz
	dim zjmonth_d
action=request("action")

select case action 
   case "zjpost"
     call zjpost
   case ""
     call main

end select	  	 


Sub zjpost()
	dim zjmonth
	zjyear=cint(request("zjyear"))
	zjmonth=cint(request("zjmonth"))
    sscj=request("sscj")
	ssbz=request("ssbz")
	url="zjqk_qtjc_post.asp?action=zjpost&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj
	
	zjmonth_d=zjmonth&"月"
	Dwt.Out "<Div style='left:6px;'>"& vbCrLf
	Dwt.Out "     <Div class='x-layOut-panel-hd'>"& vbCrLf
	Dwt.Out "        <SPAN class='x-layOut-panel-hd-text' style:'top:0px;'>"&zjyear&"年-"&zjmonth_d&" "&sscjh(sscj)&" 周检台账</span>"& vbCrLf
	Dwt.Out "     </Div>"& vbCrLf
	
if zjmonth<>0 then sql="SELECT * from sb where (year(dateadd('m',sb_test_period,sb_sczjdate))="&zjyear&" or year(sb_sczjdate)="&zjyear&") and (month(dateadd('m',sb_test_period,sb_sczjdate))="&zjmonth&"  or month(sb_sczjdate)="&zjmonth&") and sb_sscj="&sscj&" and sb_dclass=164 and sb_test_period<>0 and sb_iszj=true ORDER BY sb_id aSC "
	if zjmonth=0 then sql="SELECT * from sb where sb_sscj="&sscj&" and sb_dclass=164 and sb_iszj=true and sb_test_period<>0 ORDER BY sb_id aSC "

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then 
		message "未找到相关内容" 
	else
		Dwt.Out "<Div class='x-layOut-panel' style='WIDTH: 100%;'>"& vbCrLf
		
		Dwt.Out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.Out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.Out "     <td  class='x-td'><Div class='x-grid-hd-text'>序号</Div></td>" & vbCrLf
		Dwt.Out "      <td class='x-td'><Div class='x-grid-hd-text'>车间</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>类型</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>位号</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>装置</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>规格型号</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>测量范围</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>鉴定周期</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>计划鉴定日期</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>实际鉴定日期</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>备注</Div></td>" & vbCrLf
		Dwt.Out "      <td  class='x-td'><Div class='x-grid-hd-text'>选项</Div></td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		record=rs.recordcount
		if Trim(Request("PgSz"))="" then
		   PgSz=20
		ELSE 
		   PgSz=Trim(Request("PgSz"))
	   end if 
	   rs.PageSize = Cint(PgSz) 
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
	   rs.absolutePage = page
	   start=PgSz*Page-PgSz+1
	   rowCount = rs.PageSize
		do while not rs.eof and rowcount>0
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr class='x-grid-row x-grid-row-alt' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.Out "<tr class='x-grid-row' onmouseOut=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.Out "     <td  Class='x-td'><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
			
			  Dwt.Out "<td class='x-td' ><Div align=""center"">"&sscjh_d(rs("sb_sscj"))
              DWT.OUT "</Div></td>" & vbCrLf
			  
			  
			  
			if zclassor(rs("sb_dclass")) then 
			   if zclass(rs("sb_zclass"))="未编辑" then 
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_dclass"))&"&nbsp;</td>" & vbCrLf
			   else
			Dwt.Out "<td  class='x-td'>"&zclass(rs("sb_zclass"))&"&nbsp;</td>" & vbCrLf
			   end if 
			 end if  
			  	
			Dwt.Out "<td  class='x-td'>"&rs("sb_wh")&"&nbsp;</td>" & vbCrLf
			
			Dwt.Out "<td  class='x-td'>"&GH(rs("sb_ssGH"))&"&nbsp;</td>" & vbCrLf

			Dwt.Out "<td  class='x-td'>"&rs("sb_ggxh")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "<td  class='x-td'>"&rs("sb_c1")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'><Div align=""center"">"
			dwt.out dispalydatadict("鉴定周期",rs("sb_test_period"))
			dwt.out"&nbsp;</Div></td>" & vbCrLf

	
					dim jdzq  '检定周期判断
					dim jdinfo
					dim jdyear '检定周期换算为年
					jdzq=rs("sb_test_period")/12
					
			Dwt.Out " <td  class='x-td'><Div align=""center"">"				   
			   Dwt.out  dateadd("m",rs("sb_test_period"),rs("sb_sczjdate"))
			Dwt.out "</Div></td>" & vbCrLf
			
			dim sqlinfo,rsinfo
			dim c_text
			Dwt.Out "<td  class='x-td'><Div align=""center"">"

			
			if zjmonth<>0  then sqlinfo="SELECT * from zjinfo_qtjc where  year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("sb_id")
'			if zjmonth=0  then sqlinfo="SELECT * from zjinfo_qtjc where  dxzjyear="&zjyear&"  and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "未周检"
					c_text="<a href=zjtz_qtjc.asp?action=addzjinfo&id="&rs("sb_id")&"&sscj="&request("sscj")&"&zjdate="&zjyear&"-"&zjmonth&">完成</a>  "

			    c_text=c_text&"  <a href=zjtz_qtjc.asp?action=addzjinfo&id="&rs("sb_id")&"&sscj="&request("sscj")&">更改计划日期</a>"
			else
				 DWT.OUT RSINFO("zjdate")
				dim jdjg
				if rsinfo("zjinfo")="" then
				   jdjg="未添写鉴定结果"
				else
				   jdjg=rsinfo("zjinfo")
				end if       
				c_text="周检完成 "&jdjg
			end if 
			
			Dwt.out "</Div></td>" & vbCrLf
			Dwt.Out "      <td  class='x-td'><Div align=center>" & vbCrLf
			dwt.out c_text
			Dwt.Out "</Div></td></tr>" & vbCrLf
			c_text=""
			 RowCount=RowCount-1
	  rs.movenext
	  loop
	Dwt.Out "</table>" & vbCrLf
	   call showpage(page,url,total,record,PgSz)
   Dwt.Out "</Div>"
   end if
   Dwt.Out "</Div>"		   
   rs.close
   set rs=nothing
End Sub

'用于保存本月周检完成后所添的周检结果


sub main()
	Dwt.out "<br/><br/><br/><br/><br/>"
	dwt.out "<Div align='center'><Div class='x-dlg x-dlg-closable x-dlg-draggable x-dlg-modal' style=' WIDTH: 400px; HEIGHT: 198px'>"
	Dwt.out "  <Div class='x-dlg-hd-left'>"
	Dwt.out "    <Div class='x-dlg-hd-right'>"
	Dwt.out "      <Div class='x-dlg-hd x-unselectable'>周检设备查询</Div>"
	Dwt.out "    </Div>"
	Dwt.out "  </Div>"
	Dwt.out "  <Div class='x-dlg-dlg-body' style='WIDTH: 400px;'><Div align=left>"

	Dwt.out"<br/><form method='post' action='zjqk_qtjc_post.asp' name='form1' onsubmit='javascript:return check();'>"
	Dwt.out "<table width='100%' >"& vbCrLf
	Dwt.out"<tr><td width='20%' align='right' class='tdbg'><strong>周检月份：</strong></td> "
	Dwt.out"<td width='60%' class='tdbg'>"& vbCrLf
	Dwt.out "<select name='zjyear'>" & vbCrLf
	Dwt.out "<option value=''>选择年份</option>" & vbCrLf
	for i=year(now())-5 to year(now())+5
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=year(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
			'Dwt.out"<option value='"&i&"'>"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>年	" & vbCrLf
	Dwt.out "<select name='zjmonth'>" & vbCrLf
	Dwt.out "<option value=''>选择月份</option>" & vbCrLf
	dwt.out "<option value=0>大修</option>"
	for i=1 to 12
		Dwt.out"<option value='"&i&"'"& vbCrLf
		if i=month(now()) then Dwt.out" selected"
		Dwt.out">"&i&"</option>"& vbCrLf
	next
	Dwt.out "</select>	" & vbCrLf
	Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	Dwt.out"<strong>所属车间：</strong></td>"
	Dwt.out "<td>" & vbCrLf
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	
	Dwt.out"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 and levelid<>11"& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'"& vbCrLf
		'if session("level")=rscj("levelid") then Dwt.out "selected"
		Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
'    Dwt.out"</select>"  	 & vbCrLf
'    Dwt.out "<select name='ssbz' size='1' >" & vbCrLf
'    Dwt.out "<option  selected>选择班组分类</option>" & vbCrLf
'    Dwt.out "</select></td></tr>  "  & vbCrLf
'    Dwt.out "<script><!--" & vbCrLf
'    Dwt.out "var groups=document.form1.sscj.options.length" & vbCrLf
'    Dwt.out "var group=new Array(groups)" & vbCrLf
'    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
'    Dwt.out "group[i]=new Array()" & vbCrLf
'    Dwt.out "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
'	
'	sqlcj="SELECT * from levelname where levelclass=1  and levelid<>11"& vbCrLf
'    set rscj=server.createobject("adodb.recordset")
'    rscj.open sqlcj,conn,1,1
'    
'	do while not rscj.eof
'     ii=0		
'		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
'        set rsbz=server.createobject("adodb.recordset")
'        rsbz.open sqlbz,conn,1,1
'        if rsbz.eof and rsbz.bof then
'		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""无班组"",""0"");" & vbCrLf
'		else
'		do while not rsbz.eof
'		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
'		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
'		  ii=ii+1
'		   rsbz.movenext
'	    loop
'	    end if 
'		rsbz.close
'	    set rsbz=nothing
'
'		rscj.movenext
'	loop
'	rscj.close
'	set rscj=nothing
'    Dwt.out "var temp=document.form1.ssbz" & vbCrLf
'    Dwt.out "function redirect(x){" & vbCrLf
'    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
'    Dwt.out "temp.options[m]=null" & vbCrLf
'    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
'    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
'    Dwt.out "}" & vbCrLf
'    Dwt.out "temp.options[0].selected=true" & vbCrLf
'    Dwt.out "}//--></script" & vbCrLf
'	Dwt.out "</td></tr>" & vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='zjpost'><input  type='submit' name='Submit' value='查询' style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

	Dwt.out "  </Div></Div>"
	Dwt.out "</Div></Div>"
end sub

function zclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclass= "未编辑"
  else
     zclass=rsbody("sbclass_name")
  end if
end function

'用于分类名称显示

function zclassor(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_zclass="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclassor=false 
  else
     zclassor=true
  end if
end function



'判断是否有子分类


Call Closeconn
%>