<%@language=vbscript codepage=936 %>
<%
Option Explicit'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"%>

<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->

<%

if request("action")="lefturlchick" then 
  session("pagelevelid")=""
  
  session("pagelevelid")=request("pagelevelid")
 'message request("url")
  dwt.out "<script>location='"&replace(request("url"),"*","&")&"'</script>"
end if 

dwt.out "<html>" & vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
dwt.out "<title>系统导航菜单</title>" & vbCrLf
dwt.out "<link href='css/docs.css' rel='stylesheet' type='text/css'/>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'/>"& vbCrLf
dwt.out "<link href='css/left.css' rel='stylesheet' type='text/css'/>"& vbCrLf
dwt.out "<link href='css/dtree.css' rel='StyleSheet' type='text/css' /> "& vbCrLf
dwt.out "<script src='js/dtree.js' type='text/javascript'></script>"& vbCrLf
dwt.out " <style type='text/css'>"& vbCrLf
dwt.out " .dtree { font-family: Verdana, Geneva, Arial, Helvetica, sans-serif; font-size: 11px; white-space: nowrap;}"& vbCrLf
dwt.out " .dtree img { border: 0px; vertical-align: middle;}"& vbCrLf
dwt.out " .dtree a { text-decoration: none;}"& vbCrLf
dwt.out " .dtree a.node "& vbCrLf
dwt.out " .dtree a.nodeSel { white-space: nowrap; padding: 1px 2px 1px 2px;}"& vbCrLf
dwt.out " dtree .clip { overflow: hidden;}"& vbCrLf
dwt.out " </style>"& vbCrLf
dwt.out "</head>" & vbCrLf
dwt.out "<BODY leftmargin='0' topmargin='0' marginheight='0' marginwidth='0'>" & vbCrLf
dwt.out "<DIV class='x-layout-panel x-layout-panel-west' style='LEFT: 0px; WIDTH: 176px; '>"

dwt.out "<DIV class='x-unselectable x-layout-panel-hd x-layout-title-west' style='padding-LEFT: 10px;padding-top: 5px;WIDTH: 176px;height:20px' ><SPAN class='font:normal 14px tahoma,;' >登陆用户：" & session("username1") & "</SPAN></div>"
dwt.out "<br/>"  

'dim leftmdb,connleft,connl
dim rs,sql,leftnumb
'leftmdb="ybdata/left.mdb"
Set connleft = Server.CreateObject("ADODB.Connection")
'connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
connleft.Open connl    

sql="SELECT * from left_class where zclass=0 order by orderby asc"
set rs=server.createobject("adodb.recordset")
rs.open sql,connleft,1,1
if rs.eof and rs.bof then 
else
	do while not rs.eof
			if displaypagelevelh(session("groupid"),0,rs("id")) then 
				leftnumb=leftnumb+1
				dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
				dwt.out "  <tr>" & vbCrLf
				if rs("isbiglevel") then 
				    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='img_ext/left/left_"&leftnumb&".gif' id=menuTitle1 onclick=""window.open('left.asp?action=lefturlchick&pagelevelid="&rs("id")&"&url="&rs("url")&"','main');showsubmenu("&leftnumb&");"" style='cursor:hand;'><a href='left.asp?action=lefturlchick&pagelevelid="&rs("id")&"&url="&rs("url")&"' target='main'><span>"&rs("name")&"</span></a></td>" & vbCrLf
				else
				    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='img_ext/left/left_"&leftnumb&".gif' id=menuTitle1 onclick=""window.open('"&rs("url")&"','main');showsubmenu("&leftnumb&");"" style='cursor:hand;'><a href='"&rs("url")&"' target='main'><span>"&rs("name")&"</span></a></td>" & vbCrLf
				end if 
				dwt.out "  </tr>" & vbCrLf
				dwt.out "  <tr>" & vbCrLf
				dwt.out "    <td style='display:none' id='submenu"&leftnumb&"'>"
				dwt.out "<div class=sec_menu style='width:158'>" & vbCrLf
				dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
				dim sqlz,rsz
				if rs("id")=4 then 
					'此处有问题不能在后台自动添加,删除，因用到DTREE，随后修改
					sbmenu   
				else
					if rs("id")=125 then call dgtmenu
sqlz="SELECT * from left_class where zclass="&rs("id")&"  order by orderby asc"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,connleft,1,1
					if rsz.eof and rsz.bof then 
						dwt.out "<tr><td>无子菜单</td></tr>"
					else
						do while not rsz.eof
							if rs("isbiglevel") then '判断是否继承一级分类的权限
								if displaypagelevelh(session("groupid"),0,rs("id")) then 
									dwt.out "<tr>" & vbCrLf
									dwt.out "<td height=20 class=Glow>"
									if rsz("url")="" then 
									 dwt.out "<font color=#999999>"&rsz("name")&"</font>"
									else 
									 'dim url
									 dwt.out "<a href='left.asp?action=lefturlchick&pagelevelid="&rs("id")&"&url="&rsz("url")&"' target=main>"&rsz("name")&"</a> "
									 'if rsz("id")=13 then dwt.out "(公司控) (厂控)"
									 dwt.out "<br>" & vbCrLf
									end if 
									dwt.out "</td></tr>" & vbCrLf
								end if 
							else
								if displaypagelevelh(session("groupid"),0,rsz("id")) then 
									dwt.out "<tr>" & vbCrLf
									dwt.out "<td height=20 class=Glow>"
									if rsz("url")="" then 
									 dwt.out "<font color=#999999>"&rsz("name")&"</font>"
									else 
									 'dim url
									 dwt.out "<a href='left.asp?action=lefturlchick&pagelevelid="&rsz("id")&"&url="&rsz("url")&"' target=main>"&rsz("name")&"</a>"
									 if rsz("id")=13 then dwt.out "<br>&nbsp;&nbsp;<a href='left.asp?action=lefturlchick&pagelevelid=13&url=lsda.asp?search=gsk' target=main>公司控</a><br>&nbsp;&nbsp;<a href='left.asp?action=lefturlchick&pagelevelid=13&url=lsda.asp?search=ck' target=main>厂控</a><br>&nbsp;&nbsp;<a href='left.asp?action=lefturlchick&pagelevelid=13&url=lsda.asp?search=del' target=main>已取消</a><br>&nbsp;&nbsp;<a href='/UPLOADFILE/LSSPB.DOC' target=main>审批表</a>"
									 dwt.out "<br>" & vbCrLf
									end if 
									dwt.out "</td></tr>" & vbCrLf
								end if 
							end if 
						rsz.movenext
						loop
					end if 	
					if rs("id")=5 then call stmenu
					if rs("id")=1 then call fdmenu
					if rs("id")=2 then call zjmenu
					
					
					rsz.close
					set rsz=nothing
				end if 
				dwt.out "</table>      </div>" & vbCrLf
				dwt.out "        <div  style='width:158'>" & vbCrLf
				dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
				dwt.out "            <tr>" & vbCrLf
				dwt.out "              <td height=20></td>" & vbCrLf
				dwt.out "            </tr>" & vbCrLf
				dwt.out "          </table>" & vbCrLf
				dwt.out "      </div></td>" & vbCrLf
				dwt.out "  </tr>" & vbCrLf
				dwt.out "</table>" & vbCrLf
			end if 
	rs.movenext
	loop
end if 	
rs.close
set rs=nothing
connleft.close
set connleft=nothing


dwt.out "<SCRIPT language=javascript1.2>" & vbCrLf
dwt.out "function showsubmenu(sid){" & vbCrLf
dwt.out "    whichEl = eval('submenu' + sid);" & vbCrLf
dwt.out "    if (whichEl.style.display == 'none'){" & vbCrLf
dwt.out "        eval(""submenu"" + sid + "".style.display='';"");" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "    else{" & vbCrLf
dwt.out "        eval(""submenu"" + sid + "".style.display='none';"");" & vbCrLf
dwt.out "    }" & vbCrLf
dwt.out "for(var i=1;i<sid;i=i+1){" & vbCrLf
dwt.out "     eval(""submenu"" + i + "".style.display='none';"");" & vbCrLf
dwt.out "     }" & vbCrLf
dwt.out "for(var ii="&leftnumb&";ii>sid;ii=ii-1){" & vbCrLf
dwt.out "     eval(""submenu"" + ii + "".style.display='none';"");" & vbCrLf
dwt.out "     }" & vbCrLf
dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "</body>" & vbCrLf
dwt.out "</html>" & vbCrLf

sub sbmenu()

			dwt.out "<tr><td>"& vbCrLf
			dwt.out "<div class='dtree'>"& vbCrLf
			dwt.out "<script type='text/javascript'>"& vbCrLf
			dwt.out "	<!--"& vbCrLf
			dwt.out "	d = new dTree('d');"& vbCrLf
			dwt.out "       d.config.useCookies=false;"& vbCrLf
			dim rstree,sqltree,sb_classnumb   '根目录
			dim sb_zclassnumb  '二级
			dim rszz,sqlzz,sb_zzclassnumb '三级
			sb_classnumb=0
			'sb_zclassnumb=0
			dwt.out "d.add("&sb_classnumb&",-1,'添加新设备','left.asp?action=lefturlchick&pagelevelid=4&url=sb.asp?action=add','','main');"& vbCrLf
			'dwt.out "	d.add(2,0,'现场设备','');"& vbCrLf
			'根目录
			sqltree="SELECT * from sbclass where sbclass_zclass=0 and sbclass_isput=true order by  sbclass_orderby aSC"& vbCrLf
			set rstree=server.createobject("adodb.recordset")
			rstree.open sqltree,conn,1,1
			do while not rstree.eof

			if rstree("sbclass_id")=164 then
			sb_classnumb=sb_classnumb+1
			dwt.out "d.add("&sb_classnumb&",0,'"&rstree("sbclass_name")&"','sb_qtjc.asp?sbclassid="&rstree("sbclass_id")&"','','main');" & vbCrLf
			sb_zclassnumb=sb_classnumb
			else
			sb_classnumb=sb_classnumb+1
			dwt.out "d.add("&sb_classnumb&",0,'"&rstree("sbclass_name")&"','','','main');" & vbCrLf
			sb_zclassnumb=sb_classnumb
			end if
			
			'二级
			sqlz="SELECT * from sbclass where sbclass_zclass="&rstree("sbclass_id")&" and sbclass_isput=true order by  sbclass_orderby aSC"& vbCrLf
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conn,1,1
			if rsz.eof and rsz.bof then 
			else
			do while not rsz.eof
			sb_zclassnumb=sb_zclassnumb+1
			
			if rstree("sbclass_id")=164 then			
			dwt.out "d.add("&sb_zclassnumb&","&sb_classnumb&",'"&rsz("sbclass_name")&"','sb_qtjc.asp?sbclassid=164&sbzclassid="&rsz("sbclass_id")&"','','main');" & vbCrLf
			sb_zzclassnumb=sb_zclassnumb
			else
			
			dwt.out "d.add("&sb_zclassnumb&","&sb_classnumb&",'"&rsz("sbclass_name")&"','left.asp?action=lefturlchick&pagelevelid=4&url=sb.asp?sbclassid="&rsz("sbclass_id")&"','','main');" & vbCrLf
			sb_zzclassnumb=sb_zclassnumb
			end if
			
			'三级
			sqlzz="SELECT * from sbclass where sbclass_zclass="&rsz("sbclass_id")&" and sbclass_isput=true order by  sbclass_orderby aSC"& vbCrLf
			set rszz=server.createobject("adodb.recordset")
			rszz.open sqlzz,conn,1,1
			if rszz.eof and rszz.bof then 
			else
			do while not rszz.eof
			sb_zzclassnumb=sb_zzclassnumb+1
			dwt.out "d.add("&sb_zzclassnumb&","&sb_zclassnumb&",'"&rszz("sbclass_name")&"','sb.asp?sbclassid="&rsz("sbclass_id")&"&sbzclassid="&rszz("sbclass_id")&"','','main');" & vbCrLf
			rszz.movenext
			loop
			end if 	
			rszz.close
			set rszz=nothing
			sb_zclassnumb=sb_zzclassnumb
			rsz.movenext
			loop
			end if 	
			rsz.close
			set rsz=nothing
			sb_classnumb=sb_zclassnumb
			rstree.movenext
			loop
			rstree.close
			set rstree=nothing
			sb_zzclassnumb=sb_zzclassnumb+1
			dwt.out "		d.add("&sb_zzclassnumb&",-1,'检修记录汇总','left.asp?action=lefturlchick&pagelevelid=4&url=sb_jxjl_left.asp','','main');"& vbCrLf
			sb_zzclassnumb=sb_zzclassnumb+1
			dwt.out "		d.add("&sb_zzclassnumb&",-1,'更换记录汇总','left.asp?action=lefturlchick&pagelevelid=4&url=sb_ghjl_left.asp','','main');"& vbCrLf
			
			sb_zzclassnumb=sb_zzclassnumb+1	
			dwt.out "		d.add("&sb_zzclassnumb&",-1,'气瓶统计台帐','left.asp?action=lefturlchick&pagelevelid=4&url=qptjtz.asp','','main');"& vbCrLf
sb_zzclassnumb=sb_zzclassnumb+1	
			dwt.out "		d.add("&sb_zzclassnumb&",-1,'汇总统计','left.asp?action=lefturlchick&pagelevelid=4&url=fx/sb_fx.asp','','main');"& vbCrLf

			
			dwt.out "	document.write(d);"& vbCrLf
			
			dwt.out "	//-->"& vbCrLf
			dwt.out "</script>"& vbCrLf
			
			dwt.out "</div>"& vbCrLf
			
			dwt.out "</td></tr>"& vbCrLf

end sub

sub stmenu()
			dim rstree,sqltree
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow>"
	%>
		<div class="dtree">

	<script type="text/javascript">
		<!--

		e = new dTree('e');
        e.config.useCookies=false;

		e.add(1,-1,'试题库','left.asp?action=lefturlchick&pagelevelid=4&url=pxst.asp','','main');
		<%dim i
		i=2
		sqltree="SELECT * from pxst_class "& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,connpxjhzj,1,1
		do while not rstree.eof
			dwt.out "e.add("&i&",1,'"&rstree("class_name")&"','left.asp?action=lefturlchick&pagelevelid=5&url=pxst.asp?classid="&rstree("id")&"','','main');" & vbCrLf
		i=i+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>document.write(e);

		//-->
	</script>

</div>
<%
	
	'message session("pageleveltext")
	
    dwt.out "        </td>  </tr>" & vbCrLf

end sub
'2008_10_29增加
sub fdmenu()
			dim rstree,sqltree
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow>"
	%>
		<div class="dtree">

	<script type="text/javascript">
		<!--

		e = new dTree('e');
        e.config.useCookies=false;

		e.add(1,-1,'防冻保温','left.asp?action=lefturlchick&pagelevelid=4&url=fdbw.asp','','main');
		<%dim i
		i=2
		sqltree="SELECT * from fdbw_class "& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,connjg,1,1
		do while not rstree.eof
			dwt.out "e.add("&i&",1,'"&rstree("class_name")&"','left.asp?action=lefturlchick&pagelevelid=1&url=fdbw_wh_left.asp?classid="&rstree("id")&"','','main');" & vbCrLf
		i=i+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>document.write(e);

		//-->
	</script>

</div>
<%
	
	'message session("pageleveltext")
	
    dwt.out "        </td>  </tr>" & vbCrLf

end sub

sub zjmenu()
			dim rstree,sqltree
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow>"
	%>
		<div class="dtree">

	<script type="text/javascript">
		<!--

		e = new dTree('e');
        e.config.useCookies=false;

		e.add(1,-1,'检验测量试验设备台帐','left.asp?action=lefturlchick&pagelevelid=4&url=jycltz.asp','','main');
		e.add(2,-1,'有毒有害气体检测设备','left.asp?action=lefturlchick&pagelevelid=4&url=zjtz_qtjc.asp?sbclassid=164','','main');
		<%dim i,j
		i=3
		sqltree="SELECT * from jycl_index where zclass=1 "& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,connzj,1,1
		do while not rstree.eof
			dwt.out "e.add("&i&",1,'"&rstree("class_name")&"','left.asp?action=lefturlchick&pagelevelid=1&url=jycl_zj.asp?classid="&rstree("id")&"','','main');" & vbCrLf
		i=i+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		j=i+1
		sqltree="SELECT * from jycl_index where zclass=164"& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,connzj,1,1
		do while not rstree.eof
		dwt.out "e.add("&j&",2,'"&rstree("class_name")&"','left.asp?action=lefturlchick&pagelevelid=1&url="&rstree("url")&"&sbclassid=164&classid="&rstree("id")&"','','main');" & vbCrLf
		j=j+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>
		document.write(e);

		//-->
	</script>

</div>
<%
	
	'message session("pageleveltext")
	
    dwt.out "        </td>  </tr>" & vbCrLf

end sub

sub dgtmenu()
			dim rstree,sqltree
			dim rstree1,sqltree1,ii
			dim rstree2,sqltree2
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow>"
	%>
		<div class="dtree">

	<script type="text/javascript">
		<!--

		f = new dTree('f');
        f.config.useCookies=false;

		f.add(1,-1,'党委','left.asp?action=lefturlchick&pagelevelid=4&url=dgtzl.asp?a=0','','main');
		<%dim i
		i=2
		sqltree="SELECT * from dgtzl_index where index=0 order by orderby"& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,conndgt,1,1
		do while not rstree.eof
			dim urltmp
			
			sqltree2="SELECT * from dgtzl_index where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree2=server.createobject("adodb.recordset")
			rstree2.open sqltree2,conndgt,1,1
			if not rstree2.eof then 
				urltmp=""
			else
			    urltmp="left.asp?action=lefturlchick&pagelevelid=1&url=dgtzl.asp?indexid="&rstree("id")
			end if 
			
			
			
			
			dwt.out "f.add("&i&",1,'"&rstree("class_name")&"','"&urltmp&"','','main');" & vbCrLf
			ii=i
			sqltree1="SELECT * from dgtzl_index where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree1=server.createobject("adodb.recordset")
			rstree1.open sqltree1,conndgt,1,1
			do while not rstree1.eof
				ii=ii+1
				dim urltmp2
				if rstree1("url")<>"" then
				 urltmp2=rstree1("url")
				else
				 urltmp2= "dgtzl.asp?indexid="&rstree1("id")
				end if 
				dwt.out "f.add("&ii&","&i&",'"&rstree1("class_name")&"','left.asp?action=lefturlchick&pagelevelid=1&url="&urltmp2&"','','main');" & vbCrLf
			rstree1.movenext
			loop
			
			i=ii
			
			
			
		i=i+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>document.write(f);

		//-->
	</script>

    
    
    <script type="text/javascript">
		<!--
//工会120507
		fgh = new dTree('f');
        fgh.config.useCookies=false;

		fgh.add(1,-1,'工会','left.asp?action=lefturlchick&pagelevelid=4&url=gh.asp','','main');
		<%
		i=2
		sqltree="SELECT * from dgtzl_index_gh where index=0 order by orderby"& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,conndgt,1,1
		do while not rstree.eof
			
			
			sqltree2="SELECT * from dgtzl_index_gh where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree2=server.createobject("adodb.recordset")
			rstree2.open sqltree2,conndgt,1,1
			if not rstree2.eof then 
				urltmp=""
			else
			    urltmp="left.asp?action=lefturlchick&pagelevelid=1&url=gh.asp?indexid="&rstree("id")
			end if 
			
			
			
			
			dwt.out "fgh.add("&i&",1,'"&rstree("class_name")&"','"&urltmp&"','','main');" & vbCrLf
			ii=i
			sqltree1="SELECT * from dgtzl_index_gh where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree1=server.createobject("adodb.recordset")
			rstree1.open sqltree1,conndgt,1,1
			do while not rstree1.eof
				ii=ii+1
				
				if rstree1("url")<>"" then
				 urltmp2=rstree1("url")
				else
				 urltmp2= "gh.asp?indexid="&rstree1("id")
				end if 
				dwt.out "fgh.add("&ii&","&i&",'"&rstree1("class_name")&"','left.asp?action=lefturlchick&pagelevelid=1&url="&urltmp2&"','','main');" & vbCrLf
			rstree1.movenext
			loop
			
			i=ii
			
			
			
		i=i+1
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>document.write(fgh);

		//-->
	</script>
</div>
<%
	
	'message session("pageleveltext")
	
    dwt.out "        </td>  </tr>" & vbCrLf

end sub
%>