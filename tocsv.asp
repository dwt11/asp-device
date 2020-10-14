<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->

<%
if request("action")="fdbwmain" then call fdbwmain() '防冻保温
if request("action")="lsdamain" then call lsdamain() '联锁档案
if request("action")="pxjhmain" then call pxjhmain() '培训计划
if request("action")="pxzjmain" then call pxzjmain() '培训计划
if request("action")="dcsghmain" then call dcsghmain() 'DCS更换记录
if request("action")="dcsjxmain" then call dcsjxmain()  'DCS检修记录
if request("action")="dcssoftmain" then call dcssoftmain() ''DCS软件工作记录
if request("action")="kcgl" then call kcgl()  '库存报表  
if request("action")="zjtz"  then call zjtz() '周检台账


sub zjtz()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 

   dim sqlzjtz,rszjtz,xh,rsscdate,sqlscdate,zjmonth
   	
	'dim zjmonth
	zjyear=cint(request("zjyear"))
	zjmonth=cint(request("zjmonth"))
    sscj=request("sscj")
	ssbz=request("ssbz")
	'url="zjqk_post.asp?action=zjpost&zjyear="&zjyear&"&zjmonth="&zjmonth&"&sscj="&sscj&"&ssbz="&ssbz
	
	if zjmonth=0 then
	   zjmonth_d="大修"
	else
	   zjmonth_d=zjmonth&"月"
	end if       

	if zjmonth<>0 then sql="SELECT * from zjtz where (year(sczjdate)="&zjyear&"  or year(sczjdate)="&zjyear&"-jdzq/12) and isdx=false and month(sczjdate)="&zjmonth&" and sscj="&sscj&" and ssbz="&ssbz&" ORDER BY id aSC "
	if zjmonth=0 then sql="SELECT * from zjtz where (dxzjyear="&zjyear&"  or dxzjyear="&zjyear&"-jdzq/12) and isdx and sscj="&sscj&" and ssbz="&ssbz&" ORDER BY id aSC "

	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzj,1,1
	if rs.eof and rs.bof then 
		message "未找到相关内容" 
	else
		
		Dwt.Out "<table>"& vbCrLf
		Dwt.Out "<tr>" & vbCrLf
		Dwt.Out "     <td >序号</td>" & vbCrLf
		Dwt.Out "      <td >车间</td>" & vbCrLf
		Dwt.Out "      <td  >类型</td>" & vbCrLf
		Dwt.Out "      <td  >位号</td>" & vbCrLf
		Dwt.Out "      <td  >规格型号</td>" & vbCrLf
		Dwt.Out "      <td  >测量范围</td>" & vbCrLf
		Dwt.Out "      <td  >鉴定周期</td>" & vbCrLf
		Dwt.Out "      <td  >计划鉴定日期</td>" & vbCrLf
		Dwt.Out "      <td  >实际鉴定日期</td>" & vbCrLf
		Dwt.Out "      <td  >备注</td>" & vbCrLf
		Dwt.Out "    </tr>" & vbCrLf
		do while not rs.eof 
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.Out "<tr>"& vbCrLf
			else
			  Dwt.Out "<tr >"& vbCrLf
			end if 
			Dwt.Out "     <td  >"&xh_id&"</td>"& vbCrLf
					Dwt.Out "      <td  >"&sscjh_D(rs("sscj"))&ssbzh(rs("ssbz"))&"</td>" & vbCrLf
					Dwt.Out "      <td  >"&zjclass(rs("class"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  >"&uCase(rs("wh"))&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  >"&rs("ggxh")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  >"&rs("clfw")&"&nbsp;</td>" & vbCrLf
					Dwt.Out "      <td  >"&rs("jdzq")&"&nbsp;</td>" & vbCrLf
	
					dim jdzq  '检定周期判断
					dim jdinfo
					dim jdyear '检定周期换算为年
					jdzq=rs("jdzq")/12
					
			'上次周检日期
			Dwt.Out "      <td  ><Div align=""center"">"				   
			if rs("isdx")then 
			      if year(rs("sczjdate"))=zjyear then Dwt.out rs("dxzjyear")&"-"&"大修"
			      if year(rs("sczjdate"))<>zjyear then Dwt.out rs("dxzjyear")+jdzq&"-"&"大修"
			else
			      if year(rs("sczjdate"))=zjyear then Dwt.out rs("sczjdate")
			     
				  if year(rs("sczjdate"))<>zjyear then Dwt.out year(rs("sczjdate"))+jdzq&"-"&month(rs("sczjdate"))
			     'Dwt.out rs("sczjdate")&"sdf"&zjyear
			end if 	 	 
			Dwt.out "</td>" & vbCrLf
			 'Dwt.Out "      <td  ><Div align=""center"">"&rsscdate("zjinfo")&"</td>" & vbCrLf
			
			dim sqlinfo,rsinfo
			dim c_text
			'下次周检日期
			Dwt.Out "<td  ><Div align=""center"">"

			
			if zjmonth<>0 then sqlinfo="SELECT * from zjinfo where year(zjdate)="&zjyear&" and month(zjdate)="&zjmonth&" and zjtzid="&rs("id")
			if zjmonth=0 then sqlinfo="SELECT * from zjinfo where dxzjyear="&zjyear&" and isdx and zjtzid="&rs("id")
			set rsinfo=server.createobject("adodb.recordset")
			rsinfo.open sqlinfo,connzj,1,1
			if rsinfo.eof and rsinfo.bof then 
				dwt.out "未周检"
				'if  (year(now())>=zjyear AND month(now())>zjmonth) or (zjyear>=year(now()) AND zjmonth>month(now())) then 
					'c_text="已过期"
				'else	
				'	c_text="<a href=zjqk_post.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&">完成</a>  "
				'end if 

			    'c_text=c_text&"  <a href=zjqk_post.asp?action=complete&id="&rs("id")&"&sscj="&request("sscj")&"&ssbz="&request("ssbz")&">更改计划日期</a>"
			else
			    dwt.out rsinfo("zjdate")
				dim jdjg
				if rsinfo("zjinfo")="" then
				   jdjg="未添写鉴定结果"
				else
				   jdjg=rsinfo("zjinfo")
				end if       
				c_text="周检完成 "&jdjg
			end if 
			
			Dwt.out "</td>" & vbCrLf
			Dwt.Out "      <td  >"&rs("bz")&"&nbsp;</td>" & vbCrLf
			Dwt.Out "      </tr>" & vbCrLf
			'c_text=""
			 RowCount=RowCount-1
	  rs.movenext
	  loop
	Dwt.Out "</table>" & vbCrLf
   end if
   rs.close
   set rs=nothing

end sub

'用于分类名称显示
Function zjclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from class where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connzj,1,1
    if rsname.eof then
	'do while not rsname.eof
	else
	    zjclass=rsname("name")
		'rsname.movenext
	'loop
	end if 
	rsname.close
	set rsname=nothing
end Function










sub kcgl()

if request("zclass")=0 then  
 response.write"<Script Language=Javascript>window.alert('未选择设备分类');history.go(-1);</Script>"'
else
   call outkcgl()
end if    

end sub

sub outkcgl()
dim titlename
dim t_xc_amoney,t_sr_amoney,t_fc_amoney,t_bxc_amoney
'titlename="库存"&year(request("kcgl_date"))&"年"&month(request("kcgl_date"))&"月报表"
Response.Clear() '这一句和下面这一句很重要
response.Charset ="utf-8" '这一句很重要
'Response.ContentEncoding = System.Text.Encoding.GetEncoding("gb2312")
'Server.ScriptTimeOut = 999999
response.Buffer=true
Response.ContentType = "application/vnd.ms-excel"

Response.AddHeader "content-disposition", "inline; filename ="&year(request("kcgl_date"))&"年"&month(request("kcgl_date"))&"月"&dclass(request("zclass"))&"-"&kcclass(request("zclass"))&"报表.xls"' 

	dim sqlbb,rsbb  
	dim xh 
    sqlbb="SELECT * from kcbb where class="&request("zclass")&" and month="&month(request("kcgl_date"))
    'sqlbb="SELECT * from kcbb where class="&request("zclass")
	set rsbb=server.createobject("adodb.recordset")
    rsbb.open sqlbb,connkc,1,1
    if rsbb.eof and rsbb.bof then 
	  response.write "此月报表未生成"
	else
        	response.write "<table border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr>"
			 response.write " <td  colspan=14>"&rsbb("year")&"年"&rsbb("month")&"月</td>"
			response.write " </tr>"
			response.write " <tr>"
			response.write "  <td rowspan=2>编号</td>"
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
		 t_xc_amoney=t_xc_amoney+rsbb("xc_amoney")
		 t_sr_amoney=t_sr_amoney+rsbb("sr_amoney")
		 t_fc_amoney=t_fc_amoney+rsbb("fc_amoney")
		 t_bxc_amoney=t_bxc_amoney+rsbb("bxc_amoney")
		 rsbb.movenext
		 loop
		 	response.write " <tr >"
			response.write "  <td></td>"
			response.write "  <td>合计</td>"
			response.write "  <td></td>"
			response.write "  <td></td>"
			response.write "  <td></td>"
			response.write "  <td></td>"
			response.write "  <td>"&t_xc_amoney&"</td>"
			response.write "  <td></td>"
			response.write "  <td>"&t_sr_amoney&"</td>"
			response.write "  <td></td>"
			response.write "  <td>"&t_fc_amoney&"</td>"
			 response.write " <td></td>"
			response.write "  <td>"&t_bxc_amoney&"</td>"
			response.write "  <td>　</td>"
			response.write " </tr>"

			response.write "</table>"
	   end if
	rsbb.close
	set rsbb=nothing
end sub


'用于库存子分类名称显示
Function kcclass(classid)
	dim sqlname,rsname
	sqlname="SELECT * from kcclass where id="&classid
    set rsname=server.createobject("adodb.recordset")
    rsname.open sqlname,connkc,1,1
    if rsname.eof then
	'do while not rsname.eof
	else
	    kcclass=rsname("name")
		'rsname.movenext
	'loop
	end if 
	rsname.close
	set rsname=nothing
end Function
'用于显示父分类名称 
Function dclass(classid)
	dim sqlname,rsname
	dim sqlz,rsz
	sqlz="SELECT * from kcclass where id="&classid
    set rsz=server.createobject("adodb.recordset")
    rsz.open sqlz,connkc,1,1
    'do while not rsz.eof
	 '   kcclass=rsname("name")
		'rsname.movenext
	'loop
	   sqlname="SELECT * from class where id="&rsz("class")
       set rsname=server.createobject("adodb.recordset")
       rsname.open sqlname,connkc,1,1
       'do while not rsname.eof
	    dclass=rsname("name")
		'rsname.movenext
	'loop
	rsname.close
	set rsname=nothing
	rsz.close
	set rsz=nothing
end Function 

















sub fdbwmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 
    sqlfdbw="SELECT * from fdbw ORDER BY id DESC"
    set rsfdbw=server.createobject("adodb.recordset")
    rsfdbw.open sqlfdbw,connjg,1,1
    response.write "<table border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr>"  & vbCrLf
    response.write "<td>序号</td>" & vbCrLf
    response.write "<td>车间</td>" & vbCrLf
    response.write "<td >工号</td>" & vbCrLf
    response.write "<td >位号</td>" & vbCrLf
    response.write "<td >介质</td>" & vbCrLf
    response.write "<td>类别</td>" & vbCrLf
    response.write "<td >伴热形式</td>" & vbCrLf
    response.write "<td >投运时间</td>" & vbCrLf
    response.write "<td >备注</td>" & vbCrLf
    response.write "</tr>" & vbCrLf
    do while not rsfdbw.eof 
		select case rsfdbw("lb")
          case 1
             lb="一"
          case 2 
        	lb="二"
        end select	 
		select case rsfdbw("brxx")
          case 1
             brxx="汽"
          case 2 
        	brxx="电"
        end select	 
		xh=xh+1
                response.write "<tr  >" & vbCrLf
                response.write "     <td>"&xh&"</td>" & vbCrLf
                response.write "      <td>"&sscjh(rsfdbw("sscj"))&"</td>" & vbCrLf
                response.write "      <td >"&rsfdbw("gh")&"&nbsp;</td>" & vbCrLf
                response.write "      <td >"&rsfdbw("wh")&"&nbsp;</td>" & vbCrLf
                response.write "      <td  >"&rsfdbw("jz")&"&nbsp;</td>" & vbCrLf
                response.write "      <td>"&lb&"&nbsp;</td>" & vbCrLf
                response.write "      <td  >"&brxx&"&nbsp;</td>" & vbCrLf
	            response.write "      <td  >"&rsfdbw("date")&"&nbsp;</td>" & vbCrLf
		        response.write "      <td  >"&rsfdbw("bz")&"&nbsp;</td>" & vbCrLf
                response.write "</tr>" & vbCrLf
          rsfdbw.movenext
          loop
        response.write "</table>" & vbCrLf
end sub


sub lsdamain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 
        sqllsda="SELECT * from lsda ORDER BY lsdaid DESC"
        set rslsda=server.createobject("adodb.recordset")
        rslsda.open sqllsda,connjg,1,1
        response.write "<table border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td>序号</td><td>车间</td><td>工号</td><td>位号</td><td>分级</td>"
        response.write "<td >用途</td><td>一次件名称</td><td>单位</td><td>范围</td>"
        response.write "<td>联锁值L</td><td>联锁值H</td><td>投运状况</td><td >执行装置</td><td>备注</td></tr>"
      do while not rslsda.eof
		'xh=xh+1
                 response.write "<tr><td>"&rslsda("LSDAID")&"</td><td>"&sscjh(rslsda("sscj"))&"</td><td>"&gh(rslsda("ssgh"))&"</td>"
                response.write "<td>"&rslsda("wh")&"&nbsp;</td><td>"&rslsda("fj")&"&nbsp;</td><td>"&rslsda("yt")&"&nbsp;</td>"
                response.write "<td>"&rslsda("ycjname")&"&nbsp;</td><td>"&rslsda("cldw")&"&nbsp;</td>"
                response.write "<td>"&rslsda("clfw")&"&nbsp;</td>"
                response.write "<td>"&rslsda("lsl")&"&nbsp;</td>"
                response.write "<td>"&rslsda("lsh")&"&nbsp;</td>"
         select case rslsda("tyzk")
          case 0
             tyzk="旁路"
          case 1 
        	tyzk="投运"
        end select	 
				response.write "<td >"&tyzk&"&nbsp;</td>"
				  response.write "<td >"&rslsda("zxzz")&"&nbsp;</td>"
			    response.write "<td >"&rslsda("bz")&"&nbsp;</td>"
                response.write "</tr>"
          rslsda.movenext
          loop
        response.write "</table>"
       rslsda.close
       set rslsda=nothing
        conn.close
        set conn=nothing
end sub


sub pxjhmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&sscjh_d(request("sscj"))&request("year")&"年"&request("month")&"月"&request("titlename")&".xls"' 
      
	  '输出车间级培训计划
	  sqlpxjh="SELECT * from pxjh where ssbz=0 and sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rspxjh=server.createobject("adodb.recordset")
      rspxjh.open sqlpxjh,conne,1,1
      if rspxjh.eof and rspxjh.bof then 
        response.write "<p align='center'>未添加车间培训计划</p>" 
      else
        response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td colspan=7><div align=center>天 脊 集 团</div></td></tr><tr ><td colspan=7 ><div align=center>"&request("month")&"月份员工教育培训计划</div></td>"
        response.write "</tr><tr><td colspan=7 >"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单　位："&sscjh(request("sscj"))&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rspxjh("tbdate")&"</td>"
        response.write "</tr><tr>"
		response.write "<td ><div align=center>时间</div></td>"
        response.write "  <td ><div align=center>培训内容摘要</div></td>"
        response.write "  <td ><div align=center>培训对象及人数</div></td>"
        response.write "  <td ><div align=center>培训形式</div></td>"
        response.write "  <td ><div align=center>课时</div></td>"
        response.write "  <td ><div align=center>授课人</div></td>"
        response.write "  <td ><div align=center>备注</div></td></tr>"
        do while not rspxjh.eof
           response.write "<tr >"
           response.write "<td >"&rspxjh("month")&"."&rspxjh("day")&"</td>"
           response.write "<td >"&rspxjh("body")&"</td>"
           response.write "<td >"&rspxjh("numb")&"</td>"
           response.write "<td >"&rspxjh("xs")&"</td>"
           response.write "<td >"&rspxjh("ks")&"h</td>"
           response.write "<td>"&rspxjh("skrname")&"</td>"
           response.write "<td>"&rspxjh("bz")&"</td>"
           response.write "</tr>"
		zgname=rspxjh("zgname")
		tbrname=rspxjh("tbrname")
		   rspxjh.movenext
		loop
	 response.write "<tr>"
     response.write "<td colspan=7 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;人力资源部:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单位主管："&zgname&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
     response.write "填报人："&tbrname&"</td>"
      end if 
response.write "  </tr></table><br><br><br>"

'输出各车间所属班组培训		  
 sql="SELECT * from bzname where sscj="&request("sscj")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   do while not rs.eof
      sqlpxjh="SELECT * from pxjh where ssbz="&rs("id")&" and month="&request("month")&" and year="&request("year")
      set rspxjh=server.createobject("adodb.recordset")
      rspxjh.open sqlpxjh,conne,1,1
      if rspxjh.eof and rspxjh.bof then 
             response.write "<p align='center'>未添加"&ssbzh(rs("id"))&"培训计划</p>" 
          else
        response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td colspan=7><div align=center>天 脊 集 团</div></td></tr><tr ><td colspan=7 ><div align=center>"&request("month")&"月份员工教育培训计划</div></td>"
        response.write "</tr><tr><td colspan=7 >"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单　位："&sscjh(request("sscj"))&ssbzh(rspxjh("ssbz"))&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rspxjh("tbdate")&"</td>"
        response.write "</tr><tr>"
		response.write "<td ><div align=center>时间</div></td>"
        response.write "  <td ><div align=center>培训内容摘要</div></td>"
        response.write "  <td ><div align=center>培训对象及人数</div></td>"
        response.write "  <td ><div align=center>培训形式</div></td>"
        response.write "  <td ><div align=center>课时</div></td>"
        response.write "  <td ><div align=center>授课人</div></td>"
        response.write "  <td ><div align=center>备注</div></td></tr>"
              do while not rspxjh.eof
           response.write "<tr >"
           response.write "<td >"&rspxjh("month")&"."&rspxjh("day")&"</td>"
           response.write "<td >"&rspxjh("body")&"</td>"
           response.write "<td >"&rspxjh("numb")&"</td>"
           response.write "<td >"&rspxjh("xs")&"</td>"
           response.write "<td >"&rspxjh("ks")&"h</td>"
           response.write "<td>"&rspxjh("skrname")&"</td>"
           response.write "<td>"&rspxjh("bz")&"</td>"
           response.write "</tr>"
		zgname=rspxjh("zgname")
		tbrname=rspxjh("tbrname")
                 rspxjh.movenext
		      loop
	 response.write "<tr>"
     response.write "<td colspan=7 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;人力资源部:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单位主管："&zgname&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
     response.write "填报人："&tbrname&"</td>"
     response.write "  </tr></table><br><br><br>"
           end if 
       rs.movenext
  loop
  rs.close
  set rs=nothing
  rspxjh.close
  set rspxjh=nothing


end sub



sub pxzjmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&sscjh_d(request("sscj"))&request("year")&"年"&request("month")&"月"&request("titlename")&".xls"' 
      
	  '输出车间级培训总结
	  sqlpxzj="SELECT * from pxzj where ssbz=0 and sscj="&request("sscj")&" and month="&request("month")&" and year="&request("year")
      set rspxzj=server.createobject("adodb.recordset")
      rspxzj.open sqlpxzj,conne,1,1
      if rspxzj.eof and rspxzj.bof then 
        response.write "<p align='center'>未添加车间培训总结</p>" 
      else
        response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td colspan=7><div align=center>天 脊 集 团</div></td></tr><tr ><td colspan=7 ><div align=center>"&request("month")&"月份员工教育培训总结</div></td>"
        response.write "</tr><tr><td colspan=7 >"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单　位："&sscjh(request("sscj"))&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rspxzj("tbdate")&"</td>"
        response.write "</tr><tr>"
		response.write "<td ><div align=center>时间</div></td>"
        response.write "  <td ><div align=center>培训内容摘要</div></td>"
        response.write "  <td ><div align=center>培训对象及人数</div></td>"
        response.write "  <td ><div align=center>培训形式</div></td>"
        response.write "  <td ><div align=center>课时</div></td>"
        response.write "  <td ><div align=center>授课人</div></td>"
        response.write "  <td ><div align=center>备注</div></td></tr>"
        do while not rspxzj.eof
           response.write "<tr >"
           response.write "<td >"&rspxzj("month")&"."&rspxzj("day")&"</td>"
           response.write "<td >"&rspxzj("body")&"</td>"
           response.write "<td >"&rspxzj("numb")&"</td>"
           response.write "<td >"&rspxzj("xs")&"</td>"
           response.write "<td >"&rspxzj("ks")&"h</td>"
           response.write "<td>"&rspxzj("skrname")&"</td>"
           response.write "<td>"&rspxzj("bz")&"</td>"
           response.write "</tr>"
		zgname=rspxzj("zgname")
		tbrname=rspxzj("tbrname")
		   rspxzj.movenext
		loop
	 response.write "<tr>"
     response.write "<td colspan=7 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;人力资源部:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单位主管："&zgname&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
     response.write "填报人："&tbrname&"</td>"
      end if 
response.write "  </tr></table><br><br><br>"

'输出各车间所属班组培训		  
 sql="SELECT * from bzname where sscj="&request("sscj")
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   do while not rs.eof
      sqlpxzj="SELECT * from pxzj where ssbz="&rs("id")&" and month="&request("month")&" and year="&request("year")
      set rspxzj=server.createobject("adodb.recordset")
      rspxzj.open sqlpxzj,conne,1,1
      if rspxzj.eof and rspxzj.bof then 
             response.write "<p align='center'>未添加"&ssbzh(rs("id"))&"培训总结</p>" 
          else
        response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'><tr><td colspan=7><div align=center>天 脊 集 团</div></td></tr><tr ><td colspan=7 ><div align=center>"&request("month")&"月份员工教育培训总结</div></td>"
        response.write "</tr><tr><td colspan=7 >"
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单　位："&sscjh(request("sscj"))&ssbzh(rspxzj("ssbz"))&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rspxzj("tbdate")&"</td>"
        response.write "</tr><tr>"
		response.write "<td ><div align=center>时间</div></td>"
        response.write "  <td ><div align=center>培训内容摘要</div></td>"
        response.write "  <td ><div align=center>培训对象及人数</div></td>"
        response.write "  <td ><div align=center>培训形式</div></td>"
        response.write "  <td ><div align=center>课时</div></td>"
        response.write "  <td ><div align=center>授课人</div></td>"
        response.write "  <td ><div align=center>备注</div></td></tr>"
              do while not rspxzj.eof
           response.write "<tr >"
           response.write "<td >"&rspxzj("month")&"."&rspxzj("day")&"</td>"
           response.write "<td >"&rspxzj("body")&"</td>"
           response.write "<td >"&rspxzj("numb")&"</td>"
           response.write "<td >"&rspxzj("xs")&"</td>"
           response.write "<td >"&rspxzj("ks")&"h</td>"
           response.write "<td>"&rspxzj("skrname")&"</td>"
           response.write "<td>"&rspxzj("bz")&"</td>"
           response.write "</tr>"
		zgname=rspxzj("zgname")
		tbrname=rspxzj("tbrname")
                 rspxzj.movenext
		      loop
	 response.write "<tr>"
     response.write "<td colspan=7 >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;人力资源部:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;单位主管："&zgname&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
     response.write "填报人："&tbrname&"</td>"
     response.write "  </tr></table><br><br><br>"
           end if 
       rs.movenext
  loop
  rs.close
  set rs=nothing
  rspxzj.close
  set rspxzj=nothing


end sub

sub dcsghmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 

dim xh,sqldcsgh,rsdcsgh
sqldcsgh="SELECT * from dcsgh ORDER BY id DESC"
set rsdcsgh=server.createobject("adodb.recordset")
rsdcsgh.open sqldcsgh,conndcs,1,1
if rsdcsgh.eof and rsdcsgh.bof then 
response.write "<p align='center'>未添加DCS更换记录</p>" 
else

response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'>"
response.write "<tr>" 
response.write "     <td><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td><div align=""center""><strong>设备名称</strong></div></td>"
response.write "      <td><div align=""center""><strong>规格型号</strong></div></td>"
response.write "      <td><div align=""center""><strong>安装位置</strong></div></td>"
response.write "      <td><div align=""center""><strong>更换原因</strong></div></td>"
response.write "      <td><div align=""center""><strong>损坏时间</strong></div></td>"
response.write "      <td><div align=""center""><strong>更换时间</strong></div></td>"
response.write "      <td><div align=""center""><strong>更换人</strong></div></td>"
response.write "      <td><div align=""center""><strong>备注</strong></div></td>"
response.write "    </tr>"
           do while not rsdcsgh.eof
		xh=xh+1
                 response.write "<tr>"
                response.write "     <td ><div align=""center"">"&xh&"</div></td>"
                response.write "      <td ><div align=""center"">"&sscjh_d(rsdcsgh("sscj"))&"</div></td>"
                response.write "      <td >"&rsdcsgh("sbname")&"&nbsp;</td>"
                response.write "      <td>"&rsdcsgh("ggxh")&"&nbsp;</td>"
                response.write "      <td>"&rsdcsgh("azwz")&"&nbsp;</td>"
                response.write "      <td >"&rsdcsgh("ghyy")&"&nbsp;</td>"
                response.write "      <td >"&rsdcsgh("shdate")&"&nbsp;</td>"
                response.write "      <td>"&rsdcsgh("ghdate")&"&nbsp;</td>"
                response.write "      <td><div align=""center"">"&rsdcsgh("ghrname")&"&nbsp;</div></td>"
			    response.write "      <td>"&rsdcsgh("bz")&"&nbsp;</td>"				
                response.write "</tr>"
          rsdcsgh.movenext
          loop
        response.write "</table>"
       end if
       rsdcsgh.close
       set rsdcsgh=nothing
        conn.close
        set conn=nothing

end sub 

sub dcsjxmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 
dim xh,sqldcsjx,rsdcsjx
if request("sql1")="jxjl" then sqldcsjx="SELECT * from jxjl ORDER BY id DESC"
if request("sql1")="dcsjx" then sqldcsjx="SELECT * from dcsjx ORDER BY id DESC"
set rsdcsjx=server.createobject("adodb.recordset")
rsdcsjx.open sqldcsjx,conndcs,1,1
if rsdcsjx.eof and rsdcsjx.bof then 
response.write "<p align='center'>未添加DCS检修记录</p>" 
else

response.write "<table  border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'>"
response.write "<tr>" 
response.write "     <td><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td><div align=""center""><strong>检修原因</strong></div></td>"
response.write "      <td><div align=""center""><strong>检修内容</strong></div></td>"
response.write "      <td><div align=""center""><strong>检修人</strong></div></td>"
response.write "      <td><div align=""center""><strong>检修时间</strong></div></td>"
response.write "      <td><div align=""center""><strong>备注</strong></div></td>"
response.write "    </tr>"
           do while not rsdcsjx.eof
		xh=xh+1
                 response.write "<tr>"
                response.write "     <td><div align=""center"">"&xh&"</div></td>"
                response.write "      <td>"&sscjh(rsdcsjx("sscj"))&"</td>"
                response.write "      <td>"&rsdcsjx("jxyy")&"&nbsp;</td>"
                response.write "      <td>"&rsdcsjx("body")&"&nbsp;</td>"
                response.write "      <td><div align=""center"">"&rsdcsjx("jxrname")&"&nbsp;</div></td>"
                response.write "      <td>"&rsdcsjx("jxdate")&"&nbsp;</td>"
			    response.write "      <td>"&rsdcsjx("bz")&"&nbsp;</td>"
                response.write "</tr>"
          rsdcsjx.movenext
          loop
        response.write "</table>"
       end if
       rsdcsjx.close
       set rsdcsjx=nothing
        conn.close
        set conn=nothing

end sub

sub dcssoftmain()
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename ="&request("titlename")&".xls"' 
dim xh,sqldcssoft,rsdcssoft
sqldcssoft="SELECT * from dcssoft ORDER BY id DESC"
set rsdcssoft=server.createobject("adodb.recordset")
rsdcssoft.open sqldcssoft,conndcs,1,1
if rsdcssoft.eof and rsdcssoft.bof then 
response.write "<p align='center'>未添加DCS软件工作记录</p>" 
else

response.write "<table border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'>"
response.write "<tr>" 
response.write "     <td><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td><div align=""center""><strong>车间</strong></div></td>"
response.write "      <td ><div align=""center""><strong>作业原因</strong></div></td>"
response.write "      <td ><div align=""center""><strong>作业内容</strong></div></td>"
response.write "      <td><div align=""center""><strong>作业人</strong></div></td>"
response.write "      <td><div align=""center""><strong>作业时间</strong></div></td>"
response.write "      <td><div align=""center""><strong>备注</strong></div></td>"
response.write "    </tr>"
           do while not rsdcssoft.eof
		xh=xh+1
                 response.write "<tr>"
                response.write "     <td><div align=""center"">"&xh&"</div></td>"
                response.write "      <td>"&sscjh(rsdcssoft("sscj"))&"</td>"
                response.write "      <td >"&rsdcssoft("zyyy")&"&nbsp;</td>"
                response.write "      <td>"&rsdcssoft("body")&"&nbsp;</td>"
                response.write "      <td><div align=""center"">"&rsdcssoft("zyrname")&"&nbsp;</div></td>"
                response.write "      <td>"&rsdcssoft("zydate")&"&nbsp;</td>"
			    response.write "      <td>"&rsdcssoft("bz")&"&nbsp;</td>"
                response.write "</tr>"
          rsdcssoft.movenext
          loop
        response.write "</table>"
       end if
       rsdcssoft.close
       set rsdcssoft=nothing
        conn.close
        set conn=nothing
end sub

Call CloseConn

%> 
