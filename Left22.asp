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
dwt.out "<html>" & vbCrLf
dwt.out "<head>" & vbCrLf
dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
dwt.out "<title>系统导航菜单</title>" & vbCrLf
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
dwt.out "for(var ii=7;ii>sid;ii=ii-1){" & vbCrLf
dwt.out "     eval(""submenu"" + ii + "".style.display='none';"");" & vbCrLf
dwt.out "     }" & vbCrLf
dwt.out "}" & vbCrLf
dwt.out "</SCRIPT>" & vbCrLf
dwt.out "<link href='css/docs.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
dwt.out "<link href='css/left.css' rel='stylesheet' type='text/css'>"& vbCrLf
%>
<link href="css/dtree.css" rel="StyleSheet" type="text/css" /> 
<script src="js/dtree.js" type="text/javascript"></script>
 <style type="text/css">
 .dtree { font-family: Verdana, Geneva, Arial, Helvetica, sans-serif; font-size: 11px; white-space: nowrap;}
 .dtree img { border: 0px; vertical-align: middle;}
 .dtree a { text-decoration: none;}
 .dtree a.node, 
 .dtree a.nodeSel { white-space: nowrap; padding: 1px 2px 1px 2px;}
 dtree .clip { overflow: hidden;}
 </style>

<%
dwt.out "</head>" & vbCrLf
dwt.out "<BODY leftmargin='0' topmargin='0' marginheight='0' marginwidth='0'>" & vbCrLf
%>



<% 
dwt.out "<DIV class='x-layout-panel x-layout-panel-west' style='LEFT: 0px; WIDTH: 176px; '>"

dwt.out "<DIV class='x-unselectable x-layout-panel-hd x-layout-title-west' style='LEFT: 0px;WIDTH: 176px;' ><SPAN class=' x-layout-panel-hd-text' >当前登陆用户名：" & session("username") & "</SPAN></div>"
dwt.out "<br/>"  
    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_3.gif' id=menuTitle1 onclick='showsubmenu(1)' style='cursor:hand;'><a href='right.asp' target='main'><span>生产管理</span></a></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu1'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    	
	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href='qxtb.asp'  target='main'>分厂缺陷检查通知</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
	
	
	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href='zblog.asp'  target='main'>值班日志</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
   
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=qxdjzg.asp  target='main'>车间缺陷登记记录</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=dcsghjx.asp target='main'>DCS更换检修</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
 
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=jxjl.asp target='main'>检修记录</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
  

    'dwt.out "          <tr>" & vbCrLf
'    dwt.out "<td height=20 class=Glow><a href='zysbyz.asp'  target='main'>主要设备运转表</a><br>" & vbCrLf
'    dwt.out "        </td>  </tr>" & vbCrLf

    'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow>检修计划总结<br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf

    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf

    '不要删除此段，要用
    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_2.gif' id=menuTitle2 onclick='showsubmenu(2)' style='cursor:hand;'><a href='right.asp' target='main'><span>计量管理</span></a></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu2'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    

    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=/zjtz.asp target=main>周检台账</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=/zjqk_post.asp target=main>周检情况</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=/zjtz_class.asp target=main>分类管理</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

 	'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow>五率报表<br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf
     
 	'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow>计量器具台账<br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf

    'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow>周检率报表<br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf
    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf


  
  
   dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_1.gif' id=menuTitle3 onclick='showsubmenu(3)' style='cursor:hand;'><a href='jsgl.asp' target='main'><span>技术管理</span></a></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu3'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    
'    
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=lsda.asp?action=main target='main'>联锁档案</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=fdbw.asp target='main'>防冻保温</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
    
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=jgtz.asp target='main'>技改台账</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=tjkjgj.asp target='main'>科技信息</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
   
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=fhqc.asp target='main'>安全防护用品台账</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf



   
   
   
   
   
     dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_1.gif' id=menuTitle4 onclick='showsubmenu(4)' style='cursor:hand;'><span>设备管理</span></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu4'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    %>
	<tr><td>
	
	
	<div class="dtree">

	<script type="text/javascript">
		<!--

		d = new dTree('d');
        d.config.useCookies=false;

		d.add(0,-1,'添加新设备','sb.asp?action=add','','main');
		<%'if session("level")=0 or session("level")=7 then 
%>
		
<%'end if %>		
		d.add(2,0,'现场设备','');
		<%dim i,rs,sql,rsz,sqlz,iz
		i=3
		sql="SELECT * from sbclass where sbclass_zclass=0 and sbclass_isput=true order by  sbclass_orderby aSC"& vbCrLf
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		do while not rs.eof
			dwt.out "d.add("&i&",2,'"&rs("sbclass_name")&"','sb.asp?sbclassid="&rs("sbclass_id")&"','','main');" & vbCrLf
				iz=i
			sqlz="SELECT * from sbclass where sbclass_zclass="&rs("sbclass_id")&" and sbclass_isput=true order by  sbclass_orderby aSC"& vbCrLf
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conn,1,1
			if rsz.eof and rsz.bof then 
			else
				do while not rsz.eof
				iz=iz+1
					dwt.out "d.add("&iz&","&i&",'"&rsz("sbclass_name")&"','sb.asp?sbclassid="&rsz("sbclass_zclass")&"&sbzclassid="&rsz("sbclass_id")&"','','main');" & vbCrLf
				rsz.movenext
				loop
			end if 	
			rsz.close
			set rsz=nothing
		if iz=0 then
		  i=i+1
		else
		  i=iz+1
		end if     
		iz=0
		rs.movenext
		loop
		rs.close
		set rs=nothing


dim ii
ii=100
%>		d.add(<%=ii%>,0,'控制室设备','');
		
		d.add(<%=ii-1%>,100,'AI卡','');
		d.add(101,-1,'检修记录汇总','sb_jxjl_left.asp','','main');
		d.add(102,-1,'更换记录汇总','sb_ghjl_left.asp','','main');

		document.write(d);

		//-->
	</script>

</div>

</td></tr>
	
	
	<%
    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf

  
    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_1.gif' id=menuTitle5 onclick='showsubmenu(5)' style='cursor:hand;'><span>培训管理</span></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu5'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    
	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=pxjhzj.asp?action=pxjh target='main'>培训计划</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=pxjhzj.asp?action=pxzj target='main'>培训总结</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=pxjh_view.asp?action=addpxjh target='main'>添加培训计划</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=pxzj_view.asp?action=addpxzj target='main'>添加培训总结</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
    dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow>"
	%>
		<div class="dtree">

	<script type="text/javascript">
		<!--

		e = new dTree('e');
        e.config.useCookies=false;

		e.add(1,-1,'试题库','pxst.asp','','main');
		<%
		i=2
		sql="SELECT * from pxst_class "& vbCrLf
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connpxjhzj,1,1
		do while not rs.eof
			dwt.out "e.add("&i&",1,'"&rs("class_name")&"','pxst.asp?classid="&rs("id")&"','','main');" & vbCrLf
		i=i+1
		rs.movenext
		loop
		rs.close
		set rs=nothing
		%>document.write(e);

		//-->
	</script>

</div>
<%
	
	
	
    dwt.out "        </td>  </tr>" & vbCrLf
	

if session("level")=0 or session("level")=7 then 
	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=pxjhzj_bb.asp target='main'>报表输出</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
end if 


    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf

  
  
  
    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_1.gif' id=menuTitle6 onclick='showsubmenu(6)' style='cursor:hand;'><span>库存台账</span></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu6'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    
	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl.asp target='main'>现存</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl_sr.asp target='main'>入库信息</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

	dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl_fc.asp target='main'>出库信息</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf
	
		dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl_fcsa.asp?action=add target='main'>新入库添加</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

if session("level")=0 or session("level")=7 then 
	
dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl_bb.asp target='main'>报表输出</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

dwt.out "          <tr>" & vbCrLf
    dwt.out "<td height=20 class=Glow><a href=kcgl_class.asp target='main'>分类管理</a><br>" & vbCrLf
    dwt.out "        </td>  </tr>" & vbCrLf

end if 
    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf






   
 


    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_04.gif' id=menuTitle7 onclick='showsubmenu(7)' style='cursor:hand;'><a href='right.asp' target='main'><span class=glow>行政管理</span></a></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu7'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
			sqlz="SELECT * from xzgl_news_class"
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conna,1,1
			if rsz.eof and rsz.bof then 
			else
				do while not rsz.eof
					dwt.out "          <tr>" & vbCrLf
					dwt.out "            <td height=20 class=Glow><a href=/news.asp?classid="&rsz("id")&" target=main>"&rsz("class_name")&"</a></td>" & vbCrLf
					dwt.out "          </tr>" & vbCrLf
				rsz.movenext
				loop
			end if 	
			rsz.close
			set rsz=nothing
		
		dwt.out "          <tr>" & vbCrLf
        dwt.out "            <td height=20 class=Glow><a href=/yjhzj.asp target=main>月计划总结</a></td>" & vbCrLf
        dwt.out "          </tr>" & vbCrLf
	
	dwt.out "          <tr>" & vbCrLf
        dwt.out "            <td height=20 class=Glow><a href=/zbb.asp target=main>值班表</a></td>" & vbCrLf
        dwt.out "          </tr>" & vbCrLf
    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf





    'dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    'dwt.out "  <tr>" & vbCrLf
    'dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_05.gif' id=menuTitle7 onclick='showsubmenu(7)' style='cursor:hand;'><a href='right.asp' target='main'><span class=glow>技术资料</span></a></td>" & vbCrLf
    'dwt.out "  </tr>" & vbCrLf
    'dwt.out "  <tr>" & vbCrLf
    'dwt.out "    <td style='display:none' id='submenu7'><div class=sec_menu style='width:158'>" & vbCrLf
    'dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    
    '    dwt.out "          <tr>" & vbCrLf
    '    dwt.out "            <td height=20 class=Glow>科技成果</td>" & vbCrLf
    '    dwt.out "          </tr>" & vbCrLf
    
	'dwt.out "          <tr>" & vbCrLf
    '    dwt.out "            <td height=20 class=Glow>技术资料</td>" & vbCrLf
    '    dwt.out "          </tr>" & vbCrLf

	'dwt.out "          <tr>" & vbCrLf
     '   dwt.out "            <td height=20 class=Glow>培训讲座</td>" & vbCrLf
     '   dwt.out "          </tr>" & vbCrLf
    'dwt.out "        </table>" & vbCrLf
    'dwt.out "      </div>" & vbCrLf
    'dwt.out "        <div  style='width:158'>" & vbCrLf
   ' dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    'dwt.out "            <tr>" & vbCrLf
    'dwt.out "              <td height=20></td>" & vbCrLf
    'dwt.out "            </tr>" & vbCrLf
    'dwt.out "          </table>" & vbCrLf
    'dwt.out "      </div></td>" & vbCrLf
    'dwt.out "  </tr>" & vbCrLf
    'dwt.out "</table>" & vbCrLf


    'dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    'dwt.out "  <tr>" & vbCrLf
    'dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_1.gif' id=menuTitle103 onclick='showsubmenu(103)' style='cursor:hand;'><a href='right.asp' target='main'><span>内部邮件</span></a></td>" & vbCrLf
    'dwt.out "  </tr>" & vbCrLf
    'dwt.out "  <tr>" & vbCrLf
    'dwt.out "    <td style='display:none' id='submenu103'><div class=sec_menu style='width:158'>" & vbCrLf
    'dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
	
	'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow><a href=message.asp target='main'>收件箱</a><br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf
    
	'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow><a href=message.asp?action=add target='main'>写邮件</a><br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf

    
    
    'dwt.out "          <tr>" & vbCrLf
    'dwt.out "<td height=20 class=Glow><a href=message.asp?action=f target='main'>发件箱</a><br>" & vbCrLf
    'dwt.out "        </td>  </tr>" & vbCrLf
	   
    'dwt.out "        </table>" & vbCrLf
    'dwt.out "      </div>" & vbCrLf
    'dwt.out "        <div  style='width:158'>" & vbCrLf
    'dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    'dwt.out "            <tr>" & vbCrLf
    'dwt.out "              <td height=20></td>" & vbCrLf
    'dwt.out "            </tr>" & vbCrLf
    'dwt.out "          </table>" & vbCrLf
    'dwt.out "      </div></td>" & vbCrLf
    'dwt.out "  </tr>" & vbCrLf
    'dwt.out "</table>" & vbCrLf


    dwt.out "<table cellpadding=0 cellspacing=0 width=158 align=center>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/left_02.gif' id=menuTitle8 onclick='showsubmenu(8)' style='cursor:hand;'><a href='right.asp' target='main'>"
  if  session("levelclass")=10 then
    dwt.out "<span class=glow>后台管理</span>"
  else	 
	dwt.out "<span class=glow>用户信息管理</span>"
  end if 
	dwt.out "</a></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "  <tr>" & vbCrLf
    dwt.out "    <td style='display:none' id='submenu8'><div class=sec_menu style='width:158'>" & vbCrLf
    dwt.out "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
  if  session("levelclass")=10 then
       dwt.out  "         <tr>" & vbCrLf
	   dwt.out "         <td height=20 class=Glow><a href=userManagement.asp target=main>用户管理</a></td>" 
	   dwt.out "          </tr>" & vbCrLf


       dwt.out  "         <tr>" & vbCrLf
	   dwt.out "         <td height=20 class=Glow><a href=bzManagement.asp target=main>班组管理</a></td>" 
	   dwt.out "          </tr>" & vbCrLf
       dwt.out  "         <tr>" & vbCrLf
	   dwt.out "         <td height=20 class=Glow><a href=ghManagement.asp target=main>装置管理</a></td>" 
	   dwt.out "          </tr>" & vbCrLf
       dwt.out  "         <tr>" & vbCrLf
	   dwt.out "         <td height=20 class=Glow><a href=sb_class.asp target=main>设备管理-设备分类</a></td>" 
	   dwt.out "          </tr>" & vbCrLf
       dwt.out  "         <tr>" & vbCrLf
	   dwt.out "         <td height=20 class=Glow><a href=pxst_class.asp target=main>培训管理-试题分类</a></td>" 
	   dwt.out "          </tr>" & vbCrLf
	   
   else
        dwt.out "          <tr>" & vbCrLf
        dwt.out "            <td height=20 class=Glow><a href=usermanagement.asp?action=edit&ID="&session("userid")&" target=main>密码修改</a></td>" & vbCrLf
        dwt.out "          </tr>" & vbCrLf
  end if  
    dwt.out "        </table>" & vbCrLf
    dwt.out "      </div>" & vbCrLf
    dwt.out "        <div  style='width:158'>" & vbCrLf
    dwt.out "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    dwt.out "            <tr>" & vbCrLf
    dwt.out "              <td height=20></td>" & vbCrLf
    dwt.out "            </tr>" & vbCrLf
    dwt.out "          </table>" & vbCrLf
    dwt.out "      </div></td>" & vbCrLf
    dwt.out "  </tr>" & vbCrLf
    dwt.out "</table>" & vbCrLf

dwt.out "<div align=center><br><br><a href=bug.asp target=main><font color='#FF0000'>管理系统见意收集</font></a></div>"
if session("levelclass")=10 then dwt.out "<div align=center><br><br><a href=log.asp target=main><font color='#FF0000'>更新日志</font></a></div>"


dwt.out "<div align=center><br><font color='#ffffff'>设备管理系统<br>建议分辨率：1024 X 768</font></div>"

dwt.out "</div>"

dwt.out "</body>" & vbCrLf
dwt.out "</html>" & vbCrLf



Function cutStr(str,strlen)
    '去掉所有HTML标记<br>   
Dim re   
Set re=new RegExp  
re.IgnoreCase =True 
re.Global=True   
re.Pattern="<(.[^>]*)>"  
str=re.Replace(str,"")     
set re=Nothing   
Dim l,t,c,i  
l=Len(str)  
 t=0   
 For i=1 to l  
	 c=Abs(Asc(Mid(str,i,1)))   
	If c>255 Then    
		t=t+2    
	 Else  
		t=t+1  
	 End If    
	If t>=strlen Then  
		 cutStr=left(str,i)   
		Exit For      
	Else     
		 cutStr=str 
	 End If 
Next  
cutStr=Replace(cutStr,chr(10),"") 
cutStr=Replace(cutStr,chr(13),"") 
cutStr=Replace(cutStr,chr(32),"")
cutStr=Replace(cutStr,"【","")
cutStr=Replace(cutStr,"】","")
cutStr=Replace(cutStr,"『","")
cutStr=Replace(cutStr,"』","")


	  End Function
'rsGetAdmin.Close
'Set rsGetAdmin = Nothing
Call CloseConn
%>