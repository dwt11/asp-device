<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->
<%dim sqlghname,rsghname,ghname
dim sqlbody,rsbody

sqlbody="SELECT * from ylbbody where id="&Trim(Request("id"))
    set rsbody=server.createobject("adodb.recordset")
    rsbody.open sqlbody,conn,1,1
    if rsbody.eof and rsbody.bof then 
       response.write "<p align='center'>暂无内容</p>" 
    else
   
  
	%>
	<html>
<head>
<title>设备技术档案列表页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="22" class="topbg"><div align="center"><strong>设备一栏表</strong></div></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="tdbg2">
    <td> ★ <a href="ylb.asp?action=1">电接点压力表</a> | <a href="ylb.asp?action=3">转换器</a> | <a href="ylb.asp?action=4">调节阀附件</a> | <a href="ylb.asp?action=5" >电磁阀</a> | <a href="ylb.asp?action=6">就地调节器</a> | <a href="ylb.asp?action=7">转速探头</a> | <a href="ylb.asp?action=8">流量一次元件</a> | <a href="ylb.asp?action=9">测温一次元件</a> | <a href="ylb.asp?action=10">机组探头</a> | <a href="ylb.asp?action=11">分析</a> | <a href="ylb.asp?action=12">空调</a> | <a href="ylb.asp?action=13">皮带秤</a> | <a href="ylb.asp?action=14">调节阀</a> | <a href="ylb.asp?action=15">电动执行机构</a></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="tdbg2">
    <td> ★ 维一车间 | 维二车间 | 维三车间</td>
  </tr>
</table>
<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>
  <tr>
    <td width="708" height='22'>您现在的位置：&nbsp;技术管理&nbsp;&gt;&gt;&nbsp;设备一览表&nbsp;&gt;&gt;&nbsp;调节阀&gt;&gt;<%=rsbody("wh")%>详细内容</td>
  <td width='284' height='22' align='right'>
	  <select name='select' id='select4' onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
        <option value='' selected>按车间跳转至…</option>
        <option value='ylb.asp?action=1'>维一车间</option>
        <option value='ylb.asp?action=2'>&nbsp;&nbsp;└&nbsp;bbbbbbbb</option>
        <option value='ylb.asp?action=2'>维二车间</option>
        <option value='ylb.asp?action=2'>&nbsp;&nbsp;└&nbsp;bbbbbbbb</option>
      </select>
	  &nbsp;&nbsp;
	  <select name='JumpClass' id='JumpClass' onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
	       <option value='' selected>按分类跳转至…</option>
		   <option value='ylb.asp?action=1'>电接点压力表</option>
		   <option value='ylb.asp?action=2'>变送器</option>
		   <option value='ylb.asp?action=2'>&nbsp;&nbsp;└&nbsp;bbbbbbbb</option>
		   <option value='ylb.asp?action=3'>转换器</option>
		   <option value='ylb.asp?action=4'>调节阀附件</option>
		   <option value='ylb.asp?action=5'>电磁阀</option>
		   <option value='ylb.asp?action=6'>就地调节器</option>
		   <option value='ylb.asp?action=7'>转速探头</option>
		   <option value='ylb.asp?action=8'>流量一次元件</option>
		   <option value='ylb.asp?action=9'>测温一次元件</option>
		   <option value='ylb.asp?action=10'>机组探头</option>
  		   <option value='ylb.asp?action=11'>分析</option>
		   <option value='ylb.asp?action=12'>空调</option>
		   <option value='ylb.asp?action=13'>皮带秤</option>
		   <option value='ylb.asp?action=14'>调节阀</option>
		   <option value='ylb.asp?action=15'>电动执行机构</option>
		   
		   
    </select>	</td> 
  </tr>
</table>&nbsp;
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>位 号</strong></div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>名称</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>型号</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>流量特性</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>CV计算</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>连接方式</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>介质</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>温度</strong></div></td>
    </tr>

        
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("wh")%>&nbsp;</div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("llname")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("ggxh")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_lltx")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_cv")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_ljfs")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("gyjz")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("czwd")%>&nbsp;</div></td>
    </tr>
    
		    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>入口压力</strong></div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>△P</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>调节阀出厂编号</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>填料规格</strong></div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>填料材质</strong></div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>中封规格</strong></div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>中封材质</strong></div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>数量</strong></div></td>
    </tr>
   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("czyl")%>&nbsp;</div></td>
      <td colspan="2" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_dltp")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("ccbh")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_tlgg")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_tlcz")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("tjf_zfgg")%>&nbsp;</td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zfcz")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("shul")%>&nbsp;</div></td>
    </tr>
	
			    <tr class="title">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>异性填料</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>执行器出厂编号</strong>&nbsp;</div></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>执行器型号</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>行程转角</strong>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>弹簧压缩力</strong>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>手轮方式</strong>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>作用方式</strong>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>阀门制造厂</strong>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><strong>执行机构制造厂</strong>&nbsp;</div></td>
    </tr>
   
    <tr  class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><%=rsbody("tjf_yxtl")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxqbh")%>&nbsp;</div></td>
      <td width="14%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxqxhgg")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_xczj")%>&nbsp;</div></td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_thyl")%>&nbsp;</div></td>
      <td width="11%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_slfs")%>&nbsp;</div></td>
      <td width="9%" style="border-bottom-style: solid;border-width:1px"><%=rsbody("tjf_zyfs")%>&nbsp;</td>
      <td width="10%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_fmcj")%>&nbsp;</div></td>
      <td width="12%" style="border-bottom-style: solid;border-width:1px"><div align="center"><%=rsbody("tjf_zxjgcj")%>&nbsp;</div></td>
    </tr>

	
</table>


<table width="100%"  border="0">
  <tr class="title">
    <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center"><strong>备 注</strong></div></td>
  </tr>
  <tr class="tdbg"  onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
    <td style="border-bottom-style: solid;border-width:1px" ><%=rsbody("whbeizhu")%>&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0">
  <tr class="title">
    <td style="border-bottom-style: solid;border-width:1px" width="10%"><div align="center">检修记录&nbsp;更换记录&nbsp;编辑该内容&nbsp;删除此位号信息</div></td>
  </tr>
</table>
	<%
end if
rsbody.close
set rsbody=nothing
conn.close
set conn=nothing
%>
</body>
</html>
