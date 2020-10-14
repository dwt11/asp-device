<%@language=vbscript codepage=936 %>
<%
Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->

<%
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql

'lxclassid = Trim(Request("lxclassid"))
'if lxclassid="" then lxclassid=1
response.write "<html>"& vbCrLf
response.write "<head>" & vbCrLf
response.write "<title>库存台账管理页</title>"& vbCrLf
response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
response.write "<link href='Style.css' rel='stylesheet' type='text/css'>"& vbCrLf
response.write "</head>"& vbCrLf
response.write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
response.write"<form method='post' action='tocsv.asp' name='form1' onsubmit='javascript:return check();'>"
response.write "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>库存台账报表输出</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>选择月份：</strong></td> "
response.write"<td width='88%' class='tdbg'><script language=javascript src='/js/popcalendar.js'></script>"
response.write"<input name='kcgl_date' type='text' value="&year(now())&"-"&month(now())&" >"
response.write"<a href='#' onClick=""popUpCalendar(this,kcgl_date, 'yyyy-mm'); return false;"">"
response.write"<IMG src='/images2006/calendar/date_selector.gif' border='0' align='absMiddle'></a>  只需在需要的月份选择任意一个日期即可</td></tr>"& vbCrLf
	 response.write"<tr class='tdbg'><td width='12%' align='right' class='tdbg'>"
	 response.write"<strong>分&nbsp;&nbsp;类：</strong></td>"
     response.write "<td><select name='dclass' size='1' onChange=""redirect(this.options.selectedIndex)"">" & vbCrLf
     response.write "<option  selected value=0>选择父分类</option> "

     sql="SELECT * from class"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
    if rs.eof and rs.bof then 
          response.write"暂无分类"
      else
	  do while not rs.eof
          response.write"<option value='"&rs("id")&"'>"&rs("name")&"</option>"
	  rs.movenext
	loop
    end if 
    rs.close
    set rs=nothing
	response.write "</select>" & vbCrLf
    response.write "<select name='zclass' size='1' >" & vbCrLf
    response.write "<option  selected value=0>选择子分类</option>" & vbCrLf
    response.write "</select></td></tr>" & vbCrLf
	
	
	
	response.write "<script>" & vbCrLf
response.write "<!--" & vbCrLf


response.write "var groups=document.form1.dclass.options.length" & vbCrLf
response.write "var group=new Array(groups)" & vbCrLf
response.write "for (i=0; i<groups; i++)" & vbCrLf
response.write "group[i]=new Array()" & vbCrLf
response.write"group[0][0]=new Option(""选择子分类"",""0"");" & vbCrLf
dim sqld,rsd,rsz,sqlz
sqld="SELECT * from class"
    set rsd=server.createobject("adodb.recordset")
    rsd.open sqld,connkc,1,1
    if rsd.eof and rsd.bof then 
          response.write"暂无分类"
      else
	  do while not rsd.eof
          sqlz="SELECT * from kcclass where class="&rsd("id")
         set rsz=server.createobject("adodb.recordset")
         rsz.open sqlz,connkc,1,1
         dim ia
		 ia=0
		 if rsz.eof and rsz.bof then 
            response.write"group["&rsd("id")&"]["&ia&"]=new Option(""无子分类"","""");" & vbCrLf
         else
		 
	        do while not rsz.eof
			        response.write"group["&rsd("id")&"]["&ia&"]=new Option("""&rsz("name")&""","""&rsz("id")&""");" & vbCrLf
	        ia=ia+1
			rsz.movenext
	        loop
         end if 
         rsz.close
	  rsd.movenext
	loop
    end if 
    rsd.close
    set rsd=nothing
response.write"var temp=document.form1.zclass" & vbCrLf
response.write"function redirect(x){" & vbCrLf
response.write"for (m=temp.options.length-1;m>0;m--)" & vbCrLf
response.write"temp.options[m]=null" & vbCrLf
response.write"for (i=0;i<group[x].length;i++){" & vbCrLf
response.write"temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
response.write"}" & vbCrLf
response.write"temp.options[0].selected=true" & vbCrLf
response.write"}" & vbCrLf
response.write"//-->" & vbCrLf
response.write"</script>" & vbCrLf


response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
response.write"<input name='action' type='hidden' id='action' value='kcgl'><input name='titlename' type='hidden' id='action' value='库存'><input  type='submit' name='Submit' value='输  出' style='cursor:hand;'></td>  </tr>"
response.write"</table></form>"




response.write "<table width='80%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"& vbCrLf
response.write " <form method='post' action='kcgl_bb.asp' name='form1' ><tr class='topbg'>"& vbCrLf
response.write "   <td height='22' colspan='2' align='center'><strong>生成上月库存报表数据</strong></td>"& vbCrLf
response.write "  </tr>  "& vbCrLf
response.write"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
response.write""
	response.write"<input name='action' type='hidden' id='action' value='output'><input  type='submit' name='Submit' value='生  成' style='cursor:hand;'>"
	response.write""
	response.write"</td>  </tr>"
	response.write"<tr class='tdbg'><td>注：每月在开始新的出入库操作前请点击<生成>按钮，将上月报表生成，否则无法准确保存上月的数据</tr></td>"
	
	response.write"</form></table>"





   response.write"<div ></div>"

if request("action")="output" then call output    '生成上月报表
if request("action")="bb" then call bb    '用于报表输出,进行选择

sub output()
dim sqldclass,rsdclass '大分类
dim sqlzclass,rszclass '小分类
sqldclass="SELECT * from class"
set rsdclass=server.createobject("adodb.recordset")
rsdclass.open sqldclass,connkc,1,1
do while not rsdclass.eof
  ' response.write "<strong>"&rsdclass("name")&":</strong>&nbsp;&nbsp;&nbsp;&nbsp;"
   sqlzclass="SELECT * from kcclass where class="&rsdclass("id")
   set rszclass=server.createobject("adodb.recordset")
   rszclass.open sqlzclass,connkc,1,1
   if rszclass.bof and rszclass.eof then 
      'response.write "iiuu<BR>"
   else
      do while not rszclass.eof 
	 	 'response.write rszclass("name")&"<BR>"
		 call ymjc(rszclass("id"))'写入月未结存
		 call sr(rszclass("id"))'编辑或写入本月收入
 		 call fc(rszclass("id"))'编辑或写入本月出库
         call symjc(rszclass("id"),2007,10)'写上月未结存      
	  rszclass.movenext
      loop
   end if   
   rszclass.close
   set rszclass=nothing
   response.write "<br>"
rsdclass.movenext
loop
rsdclass.close
set rsdclass=nothing   
response.write "写入成功"& vbCrLf

end sub


sub ymjc(zclassid)'保存月未结存
       		      'response.write zclassid&"<BR>"

	     dim sqlxc,rsxc   '读取XC表中数据并保存到BB中
         sqlxc="SELECT * from xc where class="&zclassid&" and sscj=7"
         set rsxc=server.createobject("adodb.recordset")
         rsxc.open sqlxc,connkc,1,1
         if rsxc.eof and rsxc.bof then 
		      'response.write"写入成功XC1111"
		 else
			 do while not rsxc.eof
	            'response.write rsxc("name")&"<br>"
	            dim rsadd,sqladd
			    '增加上月月未结存＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
				set rsadd=server.createobject("adodb.recordset")
                sqladd="select * from kcbb" 
                rsadd.open sqladd,connkc,1,3
                rsadd.addnew
                'on error resume next
                rsadd("class")=rsxc("class")
                rsadd("wpid")=rsxc("wpid")
                rsadd("name")=rsxc("name")
                rsadd("xhgg")=rsxc("xhgg")
                rsadd("dw")=rsxc("dw")
                rsadd("dmoney")=rsxc("dmoney")
                rsadd("bxc_numb")=rsxc("numb")
                rsadd("bxc_amoney")=rsxc("amoney")
                rsadd("bz")=rsxc("bz")
	            if month(now())=1 then 
				  rsadd("year")=year(now())-1
				  rsadd("month")=12
				else
				  rsadd("year")=year(now())
				  rsadd("month")=month(now())-1
                end if 
				rsadd.update
                rsadd.close
                set rsadd=nothing
         rsxc.movenext
             loop
         end if
		 rsxc.close
         set rsxc=nothing
end sub

sub sr(zclassid)'保存收入
dim sqlsr,rssr   '读取sr表中数据并保存到BB中
        if month(now())=1 then
		   sqlsr="SELECT * from sr where class="&zclassid&" and sscj=7 and sr_year="&year(now())-1&" and sr_month=12"
		else   
		 sqlsr="SELECT * from sr where class="&zclassid&" and sscj=7 and sr_year="&year(now())&" and sr_month="&month(now())-1
        end if  
		 set rssr=server.createobject("adodb.recordset")
         rssr.open sqlsr,connkc,1,1
         if rssr.eof and rssr.bof then 
		      response.write""
		 else
			 do while not rssr.eof
	           dim sqlbb,rsbb
			    if month(now())=1 then
                  sqlbb="SELECT * from kcbb where wpid="&rssr("wpid")&" and class="&zclassid&" and year="&year(now())-1&" and month=12"
          		else   
			       sqlbb="SELECT * from kcbb where wpid="&rssr("wpid")&" and class="&zclassid&" and year="&year(now())&" and month="&month(now())-1
                end if 
			   set rsbb=server.createobject("adodb.recordset")
               rsbb.open sqlbb,connkc,1,1
               if rsbb.eof and rsbb.bof then 
	                  dim rsadd,sqladd '如果在KCBB里找不到WPID，则新加一条记录
					  set rsadd=server.createobject("adodb.recordset")
                      sqladd="select * from kcbb" 
                      rsadd.open sqladd,connkc,1,3
                      rsadd.addnew
                      'on error resume next
                      rsadd("class")=rssr("class")
                      rsadd("wpid")=rssr("wpid")
                      rsadd("name")=rssr("name")
                      rsadd("xhgg")=rssr("xhgg")
                      rsadd("dw")=rssr("dw")
                      rsadd("dmoney")=rssr("dmoney")
                      rsadd("sr_numb")=rssr("numb")
                      rsadd("sr_amoney")=rssr("amoney")
                      rsadd("bz")=rssr("bz")
	                  if month(now())=1 then 
				        rsadd("year")=year(now())-1
				        rsadd("month")=12
				      else
				        rsadd("year")=year(now())
				        rsadd("month")=month(now())-1
                      end if 
                      rsadd.update
                      rsadd.close
                      set rsadd=nothing
		       else
			          dim rsedit,sqledit'如果在KCBB里找到WPID，则编辑这一条记录的本月收入
                      set rsedit=server.createobject("adodb.recordset")
                      sqledit="select * from kcbb where id="&rsbb("id")
                      rsedit.open sqledit,connkc,1,3
                      rsedit("sr_numb")=rssr("numb")
                      rsedit("sr_amoney")=rssr("amoney")
					  rsedit.update
                      rsedit.close
                      set rsedit=nothing
	          end if
			  rsbb.close
			  set rsbb=nothing
		 rssr.movenext
		 loop	  
         end if
		 rssr.close
         set rssr=nothing
end sub


sub fc(zclassid)'保存月未出库
         dim sqlfc,rsfc   '读取fc表中数据并保存到BB中
         if month(now())=1 then 
     		 sqlfc="SELECT * from fc where class="&zclassid&" and sscj=7 and fc_year="&year(now())-1&" and fc_month=12"
		 else
		     sqlfc="SELECT * from fc where class="&zclassid&" and sscj=7 and fc_year="&year(now())&" and fc_month="&month(now())-1
         end if 
	     set rsfc=server.createobject("adodb.recordset")
         rsfc.open sqlfc,connkc,1,1
         if rsfc.eof and rsfc.bof then 
		      response.write""
		 else
			 do while not rsfc.eof
	           dim sqlbb,rsbb          
                 if month(now())=1 then 
			        sqlbb="SELECT * from kcbb where wpid="&rsfc("wpid")&" and class="&zclassid&" and year="&year(now())-1&" and month=12"
		         else
			        sqlbb="SELECT * from kcbb where wpid="&rsfc("wpid")&" and class="&zclassid&" and year="&year(now())&" and month="&month(now())-1
                 end if 
			   set rsbb=server.createobject("adodb.recordset")
               rsbb.open sqlbb,connkc,1,1
               if rsbb.eof and rsbb.bof then 
	                  dim rsadd,sqladd '如果在KCBB里找不到WPID，则新加一条记录
					  set rsadd=server.createobject("adodb.recordset")
                      sqladd="select * from kcbb" 
                      rsadd.open sqladd,connkc,1,3
                      rsadd.addnew
                      'on error resume next
                      rsadd("class")=rsfc("class")
                      rsadd("wpid")=rsfc("wpid")
                      rsadd("name")=rsfc("name")
                      rsadd("xhgg")=rsfc("xhgg")
                      rsadd("dw")=rsfc("dw")
                      rsadd("dmoney")=rsfc("dmoney")
                      rsadd("fc_numb")=rsfc("numb")
                      rsadd("fc_amoney")=rsfc("amoney")
                      rsadd("bz")=rsfc("bz")
	                  if month(now())=1 then 
				        rsadd("year")=year(now())-1
				        rsadd("month")=12
				      else
				        rsadd("year")=year(now())
				        rsadd("month")=month(now())-1
                      end if 
                      rsadd.update
                      rsadd.close
                      set rsadd=nothing
		       else
			          dim rsedit,sqledit'如果在KCBB里找到WPID，则编辑这一条记录的本月收入
                      set rsedit=server.createobject("adodb.recordset")
                      sqledit="select * from kcbb where id="&rsbb("id")
                      rsedit.open sqledit,connkc,1,3
                      rsedit("fc_numb")=rsfc("numb")
                      rsedit("fc_amoney")=rsfc("amoney")
					  rsedit.update
                      rsedit.close
                      set rsedit=nothing
	          end if
			  rsbb.close
			  set rsbb=nothing
	      rsfc.movenext
		  loop
         end if
		 rsfc.close
         set rsfc=nothing
end sub

sub symjc(zclassid,years,months)'保存上月未结存
         dim sqlfc,rsfc   '读取fc表中数据并保存到BB中&
         if month(now())=1 then 
		    sqlfc="SELECT * from kcbb where class="&zclassid&" and year="&year(now())-1&" and month=11"
		 else
		    if month(now())=2 then 
    			sqlfc="SELECT * from kcbb where class="&zclassid&" and year="&year(now())-1&" and month=12"
		    else
			    sqlfc="SELECT * from kcbb where class="&zclassid&" and year="&year(now())&" and month="&month(now())-2
			end if
		
		 end if 
         set rsfc=server.createobject("adodb.recordset")
         rsfc.open sqlfc,connkc,1,1
         if rsfc.eof and rsfc.bof then 
		      response.write""
		 else
			 do while not rsfc.eof
	           dim sqlbb,rsbb          
			    if month(now())=1 then 
			        sqlbb="SELECT * from kcbb where wpid="&rsfc("wpid")&" and year="&year(now())-1&" and month=12"
		         else
			        sqlbb="SELECT * from kcbb where wpid="&rsfc("wpid")&" and year="&year(now())&" and month="&month(now())-1
                end if 
			   set rsbb=server.createobject("adodb.recordset")
               rsbb.open sqlbb,connkc,1,1
               if rsbb.eof and rsbb.bof then 
	                  dim rsadd,sqladd '如果在KCBB里找不到WPID，则新加一条记录
					  set rsadd=server.createobject("adodb.recordset")
                      sqladd="select * from kcbb" 
                      rsadd.open sqladd,connkc,1,3
                      rsadd.addnew
                      'on error resume next
                      rsadd("class")=rsfc("class")
                      rsadd("wpid")=rsfc("wpid")
                      rsadd("name")=rsfc("name")
                      rsadd("xhgg")=rsfc("xhgg")
                      rsadd("dw")=rsfc("dw")
                      rsadd("dmoney")=rsfc("dmoney")
                      rsadd("xc_numb")=rsfc("bxc_numb")
                      rsadd("xc_amoney")=rsfc("bxc_amoney")
                      rsadd("bz")=rsfc("bz")
	                  if month(now())=1 then 
				        rsadd("year")=year(now())-1
				        rsadd("month")=12
				      else
				        rsadd("year")=year(now())
				        rsadd("month")=month(now())-1
                      end if 
                      rsadd.update
                      rsadd.close
                      set rsadd=nothing
		       else
			          dim rsedit,sqledit'如果在KCBB里找到WPID，则编辑这一条记录的本月收入
                      set rsedit=server.createobject("adodb.recordset")
                      sqledit="select * from kcbb where id="&rsbb("id")
                      rsedit.open sqledit,connkc,1,3
                      rsedit("xc_numb")=rsfc("bxc_numb")
                      rsedit("xc_amoney")=rsfc("bxc_amoney")
					  rsedit.update
                      rsedit.close
                      set rsedit=nothing
	          end if
			  rsbb.close
			  set rsbb=nothing
	      rsfc.movenext
		  loop
         end if
		 rsfc.close
         set rsfc=nothing
end sub


sub bb()
dim titlename
titlename="库存"&year(request("kcgl_bb"))&"年"&month(request("kcgl_bb"))&"月报表"
Response.AddHeader "content-disposition", "inline; filename ="&titlename&".xls"' 

	dim sqlbb,rsbb  
	dim xh 
    sqlbb="SELECT * from kcbb where class="&request("zclass")&" and month="&month(request("kcgl_date"))
    'sqlbb="SELECT * from kcbb where class="&request("zclass")
	set rsbb=server.createobject("adodb.recordset")
    rsbb.open sqlbb,connkc,1,1
    if rsbb.eof and rsbb.bof then 
	  response.write "此月报表未生成"
	else
        	response.write "<table border=1 cellpadding=0 cellspacing=0 width=""100%"">"
			response.write " <tr  >"
			 response.write " <td  colspan=14>"&rsbb("year")&"年"&rsbb("month")&"月</td>"
			response.write " </tr>"
			response.write " <tr>"
			response.write "  <td rowspan=2>序号</td>"
			response.write "  <td rowspan=2 >名称</td>"
			response.write "  <td rowspan=2 >规格</td>"
			response.write "  <td rowspan=2 >单位</td>"
			response.write "  <td rowspan=2 >单价</td>"
			response.write "  <td colspan=2 >月初库存</td>"
			response.write "  <td colspan=2 >本月收入</td>"
			response.write "  <td colspan=2 >本月发出</td>"
			response.write "  <td colspan=2 >月末结存</td>"
			response.write "  <td rowspan=2 >备注</td>"
			response.write " </tr>"
			response.write " <tr>"
			response.write "  <td>数量</td>"
			response.write "  <td>金额</td>"
			response.write "  <td>数量</td>"
			response.write "  <td>金额</td>"
			response.write "  <td>数量</td>"
			response.write "  <td>金额</td>"
			response.write "  <td>数量</td>"
			response.write "  <td>金额</td>"
			response.write " </tr>"
       do while not rsbb.eof
		xh=xh+1
			response.write " <tr >"
			response.write "  <td>"&rsbb("wpid")&"</td>"
			response.write "  <td>"&rsbb("name")&"</td>"
			response.write "  <td>"&rsbb("xhgg")&"</td>"
			response.write "  <td>"&rsbb("dw")&"</td>"
			response.write "  <td>"&rsbb("dmoney")&"</td>"
			response.write "  <td>"&rsbb("xc_numb")&"</td>"
			response.write "  <td>"&rsbb("xc_amoney")&"</td>"
			response.write "  <td>"&rsbb("sr_numb")&"</td>"
			response.write "  <td>"&rsbb("sr_amoney")&"</td>"
			response.write "  <td>"&rsbb("fc_numb")&"</td>"
			response.write "  <td>"&rsbb("fc_amoney")&"</td>"
			 response.write " <td>"&rsbb("bxc_numb")&"</td>"
			response.write "  <td>"&rsbb("bxc_amoney")&"</td>"
			response.write "  <td>　</td>"
			response.write " </tr>"
		 rsbb.movenext
		 loop
			response.write "</table>"
	   end if
	rsbb.close
	set rsbb=nothing
end sub
response.write "</body></html>"

Call CloseConn
%>