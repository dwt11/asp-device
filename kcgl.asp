<%@language=vbscript codepage=936 %>
<%
'Option Explicit
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
dim url,record,pgsz,total,page,start,rowcount,ii
dim rs,sql,title

dim pagename

Dwt.pagetop " 库存台账管理页"
if request("type")=1 then title="备件"
if request("type")=2 then title="材料"


action=request("action")
select case action
  case "add"
       if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add'添加设备分类选择
  case "edit"
       call edit
  case "saveedit"'编辑子分类
      call saveedit'编辑保存子分类
  case "saveadd"'编辑子分类
      call saveadd'编辑保存子分类
  case "fc"
       if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call fc
  case "history"
       call main
  case "savefc"'编辑子分类
      call savefc'编辑保存子分类
  case "del"
        if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
			Dwt.savesl title&"管理-"&dclass(connkc.Execute("SELECT dclass FROM xc WHERE id="&request("id"))(0))&"-"&dclass(connkc.Execute("SELECT dclass FROM xc WHERE id="&request("id"))(0)),"删除",connkc.Execute("SELECT name FROM xc WHERE id="&request("id"))(0)

			
			Set Rs = Server.CreateObject("ADODB.Recordset")
			Sql = "Delete From xc Where id="&request("id")
			Connkc.execute(Sql)
			Dwt.out "<Script Language=Javascript>history.back()</Script>"
			set rs=nothing
			set conn=nothing
		end if 
  case ""
      call main
end select	  	 


'120521保存相关三个日期
sub saveedit()    
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from xc where id="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,connkc,1,3
	on error resume next
	'rsedit("sscj")=Trim(Request("qxdj_sscj"))
	
	rsedit("dhdate")=request("dhdate")
	rsedit("jhdhdate")=request("jhdhdate")
	rsedit("sjdhdate")=request("sjdhdate")
	rsedit.update
	rsedit.close
	set rsedit=nothing
	dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub




sub edit()
	id=ReplaceBadChar(Trim(request("id")))
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from xc where id="&id
	rsedit.open sqledit,connkc,1,1
   	dwt.out"<div align=center><DIV style='WIDTH: 360px;padding-top:100px'>"& vbCrLf
	dwt.out"  <DIV class=x-box-tl>"& vbCrLf
	dwt.out"	<DIV class=x-box-tr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
	dwt.out"	</DIV>"& vbCrLf
	dwt.out"  </DIV>"& vbCrLf
	dwt.out"  <DIV class=x-box-ml>"& vbCrLf
	dwt.out"	<DIV class=x-box-mr>"& vbCrLf
	dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
	dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>编辑备件相关日期</H3>"& vbCrLf
	dwt.out"		<DIV id=form-ct>"& vbCrLf
    dwt.out"<form method='post' action='kcgl.asp' name='form1' onsubmit='javascript:return checkadd();'>"
	dwt.out"			<DIV class='x-form-ct'>"& vbCrLf

	
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px' >名称:</LABEL>"& vbCrLf
	dwt.out"				<DIV class=x-form-element style='PADDING-LEFT: 80px'>"& vbCrLf
	dwt.out"				  <INPUT class='x-form-text x-form-field' style='WIDTH: 175px' name=name value='"&rsedit("name")&"'>"& vbCrLf
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
				  


	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>定货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='dhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  value='"&rsedit("dhdate")&"'>"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>计划到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jhdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly value='"&rsedit("jhdhdate")&"' >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'>实际到货日期:</LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='sjdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly value='"&rsedit("sjdhdate")&"' >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

				  
	
	dwt.out"			</DIV>"& vbCrLf
	dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
	dwt.out"			  <DIV class='x-form-btns x-form-btns-center'>"& vbCrLf
	dwt.out"			  <input name='action' type='hidden' value='saveedit'><input name='id' type='hidden' value='"&id&"'>    <input  type='submit' name='Submit' value=' 完 成 ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
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
	dwt.out"</DIV>"& vbCrLf
	dwt.out"</div> "& vbCrLf  
	rsedit.close
	set rsedit=nothing
end sub




sub add()
	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
	
	Dwt.out "function checkamoney(){" & vbCrLf
	Dwt.out "if(document.getElementById(""checkamoney"").style.display==""none"")" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").style.display=""inline"";" & vbCrLf
			
	Dwt.out "	var szdmoney=document.getElementById(""kcgl_dmoney"").value;" & vbCrLf
	Dwt.out "	var sznumb=document.getElementById(""kcgl_numb"").value;" & vbCrLf
	Dwt.out "	if(szdmoney=="""")" & vbCrLf
	Dwt.out "	{	" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").innerHTML="" 正确输入单价能自动计算出金额!"";" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
	Dwt.out "		     return;}else" & vbCrLf
	
	Dwt.out "	      if(sznumb=="""")" & vbCrLf
	Dwt.out "	      {	" & vbCrLf
	Dwt.out "		      document.getElementById(""checkamoney"").innerHTML="" 正确输入数量能自动计算出金额!"";" & vbCrLf
	Dwt.out "		      document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
	Dwt.out "		     return;" & vbCrLf
	Dwt.out "	}" & vbCrLf
	
	Dwt.out "	var szamoney=document.getElementById(""kcgl_numb"").value*document.getElementById(""kcgl_dmoney"").value;" & vbCrLf
	
	Dwt.out "	document.getElementById(""checkamoney"").innerHTML=szamoney;" & vbCrLf
	Dwt.out "	document.getElementById(""checkamoney"").className=""ok"";" & vbCrLf
	Dwt.out "	return;" & vbCrLf
	
	Dwt.out "    }" & vbCrLf
	Dwt.out "</SCRIPT>" & vbCrLf
	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
	Dwt.out "function checkadd(){" & vbCrLf
	Dwt.formcheck "form1","kcgl_dclass","一级分类未选择",0
	Dwt.formcheck "form1","kcgl_zclass","二级分类未选择",0
	Dwt.formcheck "form1","kcgl_name","名称未添写",0
	Dwt.formcheck "form1","kcgl_xhgg","规格型号未添写",0
	Dwt.formcheck "form1","kcgl_dw","单位未添写",0
	Dwt.formcheck "form1","kcgl_dmoney","单价未添写",0
	Dwt.formcheck "form1","kcgl_numb","数量未添写",0
	Dwt.formcheck "form1","kcgl_date","入库时间未添写",0
	Dwt.out "}" & vbCrLf
	Dwt.out "</SCRIPT>" & vbCrLf

	Dwt.lable_title "kcgl.asp","form1",title&"-入库添加","checkadd" 
	Dwt.lable_input "所属车间","kcgl_sscj",1000,sscjh(1000),true,false,""
	Dwt.lable_input "操作人","kcgl_userid",session("userid"),session("username1"),true,false,""
	dim rknumb
	Randomize timer
	dim rktext
	if request("type")=1 then rktext="BJ"
	if request("type")=2 then rktext="CL"
	rknumb=rktext&"RK"&year(now())&month(now())&day(now())&hour(now())&minute(now())&int(Rnd*(second(now())*100))
	'Dwt.lable_input "入库单号","kcgl_rknumb",rknumb,rknumb,true,false,""
	'Dwt.out "入库单号:"&rknumb
	Dwt.out"<Div class='x-form-item'>"& vbCrLf
	Dwt.out"<LABEL style='WIDTH: 75px'><Div align=right>入库单号:</Div></LABEL>"& vbCrLf
	Dwt.out"<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out "<span style='font-size:14px'>"&rknumb&"</span>"
	Dwt.out "<input type='hidden' name='kcgl_rknumb' value='"&rknumb&"'>"
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"</Div>"& vbCrLf
	Dwt.out"<Div class=x-form-clear-left></Div>"& vbCrLf 
 
     dim rscj,sqlcj
 
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>分类:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
	Dwt.out "<select name='kcgl_dclass' size='1' id='cat1' onChange=""selectpc(this.value,'b',document.form1.kcgl_zclass)"">"& vbCrLf
	Dwt.out "  <option selected value='0'>选择一级分类</option>"& vbCrLf
	sql="SELECT * from class where dclass=0 and type="&request("type")
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
    do while not rs.eof
       	Dwt.out"<option value='"&rs("id")&"'>"&rs("name")&"</option>"& vbCrLf
		rs.movenext
	loop
	rs.close
	set rs=nothing
	Dwt.out "</select>"& vbCrLf
	Dwt.out "<select name='kcgl_zclass' size='1' id='cat2' >"& vbCrLf
	Dwt.out "  <option selected value='0'>选择二级分类</option>"& vbCrLf
	Dwt.out "</select>"& vbCrLf
	Dwt.out "<script language='javascript'>"& vbCrLf
	Dwt.out "function selectpc(parentValue,child,addObj){"& vbCrLf
    dim b,bv,b_p,sqlz,rsz
	sql="SELECT * from class where dclass=0 "& vbCrLf
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
	 b="var b =   new Array("
	bv="var bv =   new Array("
	b_p="var b_p =   new Array("
   
	do while not rs.eof
		sqlz="SELECT * from class where dclass="&rs("id")
        set rsz=server.createobject("adodb.recordset")
        rsz.open sqlz,connkc,1,1
        if rsz.eof and rsz.bof then
		   b=b&"'无二级分类',"
		   bv=bv&"'0',"
		   b_p=b_p&"'"&rs("id")&"',"
		else
		do while not rsz.eof
			b=b&"'"&rsz("name")&"',"
			bv=bv&"'"&rsz("id")&"',"
			b_p=b_p&"'"&rs("id")&"',"
		   rsz.movenext
	    loop
	    end if 
		rsz.close
	    set rsz=nothing
		rs.movenext
	loop
	rs.close
	set rs=nothing
	b=left(b,len(b)-1)
	bv=left(bv,len(bv)-1)
	b_p=left(b_p,len(b_p)-1)
	b=b&");"
	bv=bv&");"
	b_p=b_p&");"
	Dwt.out b & vbCrLf
	Dwt.out bv & vbCrLf
	Dwt.out b_p & vbCrLf
	Dwt.out "var labelValue = new Array();"& vbCrLf
	Dwt.out "var labelText =  new Array();"& vbCrLf
	Dwt.out "var k = 0;"& vbCrLf
	Dwt.out "cObj = eval(child);"& vbCrLf
	Dwt.out "cObjV = eval(child+'v');"& vbCrLf
	Dwt.out "cpObj = eval(child + '_p');"& vbCrLf
	Dwt.out "for(i=0; i<cpObj.length; i++)"& vbCrLf
	Dwt.out "{"& vbCrLf
	Dwt.out "	if(cpObj[i] == parentValue)"& vbCrLf
	Dwt.out "	{"& vbCrLf
	Dwt.out "		labelText[k] =  cObj[i];"& vbCrLf
	Dwt.out "		labelValue[k] =	cObjV[i]; "& vbCrLf
	Dwt.out "		k++;"& vbCrLf
	Dwt.out "	}"& vbCrLf
	Dwt.out "}"& vbCrLf
	Dwt.out "addObj.options.length = 0;"& vbCrLf
	Dwt.out "addObj.options[0] = new Option('选择二级分类','0');"& vbCrLf
	Dwt.out "for(i = 0; i < labelText.length; i++) {"& vbCrLf
	Dwt.out "	addObj.add(document.createElement('option'));"& vbCrLf
	Dwt.out "	addObj.options[i+1].text=labelText[i];"& vbCrLf
	Dwt.out "	addObj.options[i+1].value=labelValue[i];"& vbCrLf
	Dwt.out "}"& vbCrLf
	Dwt.out "addObj.selectedIndex = 0;"& vbCrLf
    Dwt.out "}"& vbCrLf
    Dwt.out "</script>"& vbCrLf
	Dwt.out" <span class='tips'>选择一级分类后选择二级分类</span>"& vbCrLf
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
		Dwt.complete_a
    Dwt.lable_input_complete "名称","kcgl_name",true,"输入空格显示已存数据","kcgl","name","xc"
    Dwt.lable_input_complete "规格型号","kcgl_xhgg",true,"输入空格显示已存数据","kcgl","xhgg","xc"
	Dwt.lable_input "单位","kcgl_dw","","",false,true,""

	
	
   
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>单价:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<input type='text' name='kcgl_dmoney' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;""   onBlur=""checkamoney()"" > 元 "   
    Dwt.out " <span class='red'>*</span>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>数量:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<input type='text' name='kcgl_numb' onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" onBlur=""checkamoney()"" > "   
    Dwt.out " <span class='red'>*</span>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>金额:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<Div id=""checkamoney"" style=""display:none"" class=""ok""></Div>元"   
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
		
		
	Dwt.out"<input name='type' type='hidden' value='"&request("type")&"'>"	
	Dwt.lable_input_date "入库日期","kcgl_date",date(),false,true,""
	Dwt.lable_input "存放地址","kcgl_adress","","",false,false,""
	
	
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'><Div align=right>定货日期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='dhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'><Div align=right>计划到货日期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='jhdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
	dwt.out"				<LABEL style='WIDTH: 85px'><Div align=right>实际到货日期:</div></LABEL>"& vbCrLf
	dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    dwt.out"<input name='sjdhdate' style='WIDTH: 175px' onClick='new Calendar(0).show(this)' readOnly  >"
	dwt.out"				</DIV>"& vbCrLf
	dwt.out"			  </DIV>"& vbCrLf
	dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf

	
	
	Dwt.lable_input "备注","kcgl_bz","","",false,false,""
	Dwt.lable_footer "saveadd"," 添 加 ",false,"","" 


end sub


sub saveadd()    
	  '保存到显存表中
	  dim rsadd,sqladd
	  dim sscj
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from xc" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
      'sscj=request("kcgl_sscj")
		  'if sscj="" then sscj=7
       dim xcid
	  xcid=rsadd("id")
   	  rsadd("wpid")=0
	  rsadd("type")=request("type")
	  rsadd("dclass")=request("kcgl_dclass")
	  rsadd("zclass")=request("kcgl_zclass")
	  rsadd("sscj")=request("kcgl_sscj")
      'on error resume next
      rsadd("name")=Trim(request("kcgl_name"))
      rsadd("xhgg")=request("kcgl_xhgg")
      rsadd("dw")=request("kcgl_dw")
      rsadd("dmoney")=request("kcgl_dmoney")
      rsadd("numb")=request("kcgl_numb")
      rsadd("amoney")=request("kcgl_dmoney")*request("kcgl_numb")
      rsadd("bz")=request("kcgl_bz")
	  rsadd("rcdate")=request("kcgl_date")
	  rsadd("userid")=request("kcgl_userid")
	  rsadd("crknumb")=request("kcgl_rknumb")
		  rsadd("adress")=request("kcgl_adress")
	on error resume next
 	rsadd("dhdate")=request("dhdate")
	rsadd("jhdhdate")=request("jhdhdate")
	rsadd("sjdhdate")=request("sjdhdate")
     rsadd.update
      rsadd.close
      set rsadd=nothing
	  
	  	  '保存到历史表中
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from history" 
      rsadd.open sqladd,connkc,1,3
      rsadd.addnew
   	  rsadd("wpid")=0
	  rsadd("type")=request("type")
      rsadd("dclass")=request("kcgl_dclass")
      rsadd("zclass")=request("kcgl_zclass")
      rsadd("sscj")=request("kcgl_sscj")
      on error resume next
      rsadd("lytxt")=request("kcgl_lytxt")
	  rsadd("name")=Trim(request("kcgl_name"))
      rsadd("xhgg")=request("kcgl_xhgg")
      rsadd("dw")=request("kcgl_dw")
      rsadd("dmoney")=request("kcgl_dmoney")
      rsadd("numb")=request("kcgl_numb")
      rsadd("amoney")=request("kcgl_dmoney")*request("kcgl_numb")
  	  rsadd("rcdate")=request("kcgl_date")
	  rsadd("bz")=request("kcgl_bz")
	  rsadd("userid")=request("kcgl_userid")
	  rsadd("crknumb")=request("kcgl_rknumb")
	  rsadd("adress")=request("kcgl_adress")
 	
      rsadd.update
      rsadd.close
      set rsadd=nothing
      
	  dim titlename
	  if request("type")=1 then 
	    titlename="备件管理"
      else
	    titlename="材料管理"
	  end if 	
	  Dwt.savesl titlename&"-"&dclass(request("kcgl_dclass"))&"-"&dclass(request("kcgl_zclass")),"入库添加",Trim(request("kcgl_name"))&" 数量："&request("kcgl_rknumb")

	  Dwt.out"<Script Language=Javascript>location.href='kcgl.asp?type="&request("type")&"';</Script>"
end sub


sub fc()
	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
	
	Dwt.out "function checkamoney(){" & vbCrLf
	Dwt.out "if(document.getElementById(""checkamoney"").style.display==""none"")" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").style.display=""inline"";" & vbCrLf
			
	Dwt.out "	var szdmoney=document.getElementById(""kcgl_dmoney"").value;" & vbCrLf
	Dwt.out "	var sznumb=document.getElementById(""kcgl_numb"").value;" & vbCrLf
	Dwt.out "	if(szdmoney=="""")" & vbCrLf
	Dwt.out "	{	" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").innerHTML="" 正确输入单价能自动计算出金额!"";" & vbCrLf
	Dwt.out "		document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
	Dwt.out "		     return;}else" & vbCrLf
	
	Dwt.out "	      if(sznumb=="""")" & vbCrLf
	Dwt.out "	      {	" & vbCrLf
	Dwt.out "		      document.getElementById(""checkamoney"").innerHTML="" 正确输入数量能自动计算出金额!"";" & vbCrLf
	Dwt.out "		      document.getElementById(""checkamoney"").className=""error"";" & vbCrLf
	Dwt.out "		     return;" & vbCrLf
	Dwt.out "	}" & vbCrLf
	
	Dwt.out "	var szamoney=document.getElementById(""kcgl_numb"").value*document.getElementById(""kcgl_dmoney"").value;" & vbCrLf
	
	Dwt.out "	document.getElementById(""checkamoney"").innerHTML=szamoney;" & vbCrLf
	Dwt.out "	document.getElementById(""checkamoney"").className=""ok"";" & vbCrLf
	Dwt.out "	return;" & vbCrLf
	
	Dwt.out "    }" & vbCrLf
	Dwt.out "</SCRIPT>" & vbCrLf
	 Dwt.out "<SCRIPT language=javascript>" & vbCrLf
	Dwt.out "function checkfc(){" & vbCrLf
	Dwt.formcheck "form1","kcgl_qxtxt","出库对象未选择",0
	Dwt.formcheck "form1","kcgl_numb","数量未添写",0
	Dwt.formcheck "form1","kcgl_date","出库日期未添写",0
	Dwt.out "  if(1*document.form1.kcgl_xynumb.value<1*document.form1.kcgl_numb.value){" & vbCrLf
	Dwt.out "      alert('出库数量大于现有数量！');" & vbCrLf
	Dwt.out "  document.form1.kcgl_numb.focus();" & vbCrLf
	Dwt.out "      return false;" & vbCrLf
	Dwt.out "    }" & vbCrLf
	Dwt.out "    }" & vbCrLf
	  Dwt.out"</Script>"
  dim id 
   dim rscj,sqlcj
   dim classname
   dim rsfc,sqlfc
   
   id=ReplaceBadChar(Trim(request("id")))
   set rsfc=server.createobject("adodb.recordset")
   sqlfc="select * from xc where id="&id
   rsfc.open sqlfc,connkc,1,1
   
   
    Dwt.lable_title "kcgl.asp","form1",title&"-出库添写","checkfc" 
	Dwt.lable_input "所属车间","kcgl_sscj",rsfc("sscj"),sscjh(rsfc("sscj")),true,false,""

  
  
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>出库对象:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
   Dwt.out"<select id=""x"" name='kcgl_qxtxt' size='1'  onchange=""edit1(this, this.getElementsByTagName('option')[selectedIndex].innerText);"">"
   Dwt.out"<option >请选择出库对象</option>"
   if rsfc("sscj")=1000 then 
   sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
   set rscj=server.createobject("adodb.recordset")
   rscj.open sqlcj,conn,1,1
   do while not rscj.eof
       	'出库对象下拉列表中不显示已属于的车间
		if rscj("levelid")=rsfc("sscj") then 
		  Dwt.out""
		else
		   Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    end if 
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
	end if 
	if rsfc("sscj")<8 then 
	  Dwt.out "<option value=1000>现场使用</option>"& vbCrLf
	  Dwt.out "<option >自定义</option>"& vbCrLf
	end if   
    Dwt.out"</select>"  	& vbCrLf 
	Dwt.out "<input type='text' id=""s"" style='display:none' onblur=""edit2(this, value)"" />"& vbCrLf
	'下面段JAVASCRIPT用于自定义SELECT
	Dwt.out "<script language=""JavaScript"">"& vbCrLf
	Dwt.out "	function edit1(obj, str){"& vbCrLf
	Dwt.out "		if(str == ""自定义""){"& vbCrLf
	Dwt.out "			obj.style.display = ""none"";"& vbCrLf
	Dwt.out "			form1.s.style.display = """";"& vbCrLf
	Dwt.out "		}"& vbCrLf
	Dwt.out "	}"& vbCrLf
	Dwt.out "	function edit2(obj, str){"& vbCrLf
	Dwt.out "		var d = document.createElement(""option"");"& vbCrLf
	Dwt.out "		d.value = str;"& vbCrLf
	Dwt.out "		d.innerText = str;"& vbCrLf
	Dwt.out "		d.selected = ""true"";"& vbCrLf
	Dwt.out "		form1.x.appendChild(d);"& vbCrLf
	Dwt.out "		obj.style.display = ""none"";"& vbCrLf
	Dwt.out "		form1.x.style.display = """";"& vbCrLf
	Dwt.out "	}"& vbCrLf
	Dwt.out "	</script>"& vbCrLf
	
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.lable_input "操作人","kcgl_userid",session("userid"),session("username1"),true,false,""
    if rsfc("sscj")=1000 then 
		dim cknumb
		Randomize timer
		cknumb="BJCK"&year(now())&month(now())&day(now())&hour(now())&minute(now())&int(Rnd*second(now()))
		Dwt.lable_input "出库单号","kcgl_cknumb",cknumb,cknumb,true,false,""
    end if 
	Dwt.lable_input "分类","","",dclass(rsfc("dclass"))&"-"&dclass(rsfc("zclass")),true,false,""
	Dwt.lable_input "名称","","",rsfc("name"),true,false,""
	Dwt.lable_input "规格型号","","",rsfc("xhgg"),true,false,""
	Dwt.lable_input "单位","","",rsfc("dw"),true,false,""
   
	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>单价:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<input type='text' disabled='disabled' value="&rsfc("dmoney")&">元"   
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	Dwt.out"			  <Div class='x-form-item'>"& vbCrLf
	Dwt.out"				<LABEL style='WIDTH: 75px'><Div align=right>现有数量:</Div></LABEL>"& vbCrLf
	Dwt.out"				<Div class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
    Dwt.out"<input type='text' name='kcgl_xynumb' disabled='disabled' value="&rsfc("numb")&" >&nbsp;&nbsp;&nbsp;&nbsp;出库数量:<input type='text' name='kcgl_numb' value="&rsfc("numb")&" > "   
    Dwt.out " <span class='red'>*</span>"
	Dwt.out"				</Div>"& vbCrLf
	Dwt.out"			  </Div>"& vbCrLf
	Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf

	 	Dwt.out "<input type='hidden' name='type' value='"&request("type")&"'>"
	Dwt.lable_input_date "出库日期","kcgl_date",date(),false,true,""
	Dwt.lable_input "存放地址","kcgl_adress","","",false,false,""
	Dwt.lable_input "备注","kcgl_bz","","",false,false,""
	Dwt.lable_footer "savefc"," 完 成 ",true,"id",id 
    rsfc.close
    set rsfc=nothing
end sub



sub savefc()    '保存出存信息
	  '保存到出库表中
  dim rsfc,sqlfc
  dim sscj
  dim rscheck,sqlcheck
 '将出库信息保存到XC所属车间保存为出库对象
  set rscheck=server.createobject("adodb.recordset")
  sqlcheck="select * from xc where id="&request("id")
  rscheck.open sqlcheck,connkc,1,3
   '如果是 出库,则新建所对应出库对象车间的XC数据,并在history中新建数据.
   '否则修改:车间出库,对应车间出库数量,并在history中新建出库对象的数据
   
   
      if rscheck("sscj")=1000 then 
		  dim rsadd,sqladd
		 ' dim sscj
		  set rsadd=server.createobject("adodb.recordset")
		  sqladd="select * from xc" 
		  rsadd.open sqladd,connkc,1,3
		  rsadd.addnew
		  'sscj=request("kcgl_sscj")
			  'if sscj="" then sscj=7
		   dim xcid
		  xcid=rscheck("id")
		  rsadd("wpid")=xcid
		  rsadd("dclass")=rscheck("dclass")
		  rsadd("zclass")=rscheck("zclass")
		  rsadd("sscj")=request("kcgl_qxtxt")
		  on error resume next
		  rsadd("name")=rscheck("name")
		  rsadd("xhgg")=rscheck("xhgg")
		  rsadd("dw")=rscheck("dw")
		  rsadd("dmoney")=rscheck("dmoney")
		  rsadd("numb")=request("kcgl_numb")
		  rsadd("amoney")=rscheck("dmoney")*request("kcgl_numb")
		  rsadd("bz")=request("kcgl_bz")
		  rsadd("rcdate")=request("kcgl_date")
		  rsadd("userid")=request("kcgl_userid")
		  rsadd("crknumb")=request("kcgl_cknumb")
		  rsadd("adress")=request("kcgl_adress")
 		  rsadd("type")=request("type")
		  rsadd.update
		  rsadd.close
		  set rsadd=nothing
     
		 
		  set rsadd=server.createobject("adodb.recordset")
		  sqladd="select * from history" 
		  rsadd.open sqladd,connkc,1,3
		  rsadd.addnew
		  'sscj=request("kcgl_sscj")
			  'if sscj="" then sscj=7
		  ' dim xcid
		  xcid=rscheck("id")
		  rsadd("wpid")=xcid
		  rsadd("dclass")=rscheck("dclass")
		  rsadd("zclass")=rscheck("zclass")
		  rsadd("sscj")=request("kcgl_qxtxt")
		 ' on error resume next
		  rsadd("name")=rscheck("name")
		  rsadd("xhgg")=rscheck("xhgg")
		  rsadd("dw")=rscheck("dw")
		  rsadd("dmoney")=rscheck("dmoney")
		  rsadd("numb")=request("kcgl_numb")
		  rsadd("amoney")=rscheck("dmoney")*request("kcgl_numb")
		  rsadd("bz")=request("kcgl_bz")
		  rsadd("rcdate")=request("kcgl_date")
		  rsadd("userid")=request("kcgl_userid")
		  rsadd("crknumb")=request("kcgl_cknumb")
		  rsadd("adress")=request("kcgl_adress")
 		  rsadd("type")=request("type")
		  rsadd.update
		  rsadd.close
		  set rsadd=nothing


	 else
		  set rsadd=server.createobject("adodb.recordset")
		  sqladd="select * from history" 
		  rsadd.open sqladd,connkc,1,3
		  rsadd.addnew
		  'sscj=request("kcgl_sscj")
			  'if sscj="" then sscj=7
		   'dim xcid
		  xcid=rscheck("id")
		  rsadd("wpid")=xcid
		  rsadd("dclass")=rscheck("dclass")
		  rsadd("zclass")=rscheck("zclass")
		  rsadd("sscj")=request("kcgl_sscj")
		  'on error resume next
		  rsadd("name")=rscheck("name")
		  rsadd("xhgg")=rscheck("xhgg")
		  rsadd("dw")=rscheck("dw")
		  rsadd("dmoney")=rscheck("dmoney")
		  rsadd("numb")=request("kcgl_numb")
		  rsadd("amoney")=rscheck("dmoney")*request("kcgl_numb")
		  rsadd("bz")=request("kcgl_bz")
		  rsadd("rcdate")=request("kcgl_date")
		  rsadd("qx")=request("kcgl_qxtxt")
		  rsadd("userid")=request("kcgl_userid")
		  rsadd("crknumb")=request("kcgl_cknumb")
		  rsadd("adress")=request("kcgl_adress")
 		  rsadd("type")=request("type")
		  rsadd.update
		  rsadd.close
		  set rsadd=nothing
	 end if 
  rscheck.close
  set rscheck=nothing


	'如果出库数量=现有数量刚删除源数据,否则将现有数量减去出库数量,修改保存.. 如果等于这个功能暂不用
'	 if rscheck("numb")=Cint(request("kcgl_numb")) then 
'	    dim rsdel,sqldel
'	    set rsdel=server.createobject("adodb.recordset")
'        sqldel="delete * from xc where id="&request("id")
'       rsdel.open sqldel,connkc,1,3
'     else
	  dim rsedit,sqledit
	  set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from xc where id="&request("id")
      rsedit.open sqledit,connkc,1,3
      rsedit("numb")=rsedit("numb")-request("kcgl_numb")
      rsedit("amoney")=request("kcgl_dmoney")*rsedit("numb")
      'rsedit("rcdate")=request("kcgl_fcdate")
	  rsedit.update
      rsedit.close
      set rsedit=nothing
	 
	'end if   
	
	  
  Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
  'end if 
end sub

sub main()
	Dwt.out "<SCRIPT language=javascript1.2>" & vbCrLf
	Dwt.out "function showsubmenu(sid){" & vbCrLf
	Dwt.out "      	 var ss='xxx'+sid;" & vbCrLf
	Dwt.out "    whichEl = eval('info' + sid);" & vbCrLf
	Dwt.out "    if (whichEl.style.display == 'none'){" & vbCrLf
	Dwt.out "        eval(""info"" + sid + "".style.display='block';"");" & vbCrLf
	Dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i6.gif' />"";" & vbCrLf
	Dwt.out "    }" & vbCrLf
	Dwt.out "    else{" & vbCrLf
	Dwt.out "        eval(""info"" + sid + "".style.display='none';"");" & vbCrLf
	Dwt.out "        document.getElementById(ss).innerHTML=""<img src='/img_ext/i7.gif' />"";" & vbCrLf
	Dwt.out "    }" & vbCrLf
	Dwt.out "}" & vbCrLf
	Dwt.out "</SCRIPT>" & vbCrLf


			dim totalamoney '合计页里的总金额
	
	dim sqlbody,rsbody,xh
	
	if request("class")="" and request("sscj")="" or request("sscj")=1000 then 
	   url="kcgl.asp?type="&request("type")
	   sqlbody="SELECT * from xc where wpid=0 and type="&request("type")
       pagename="分厂现存"
	end if 
	if request("keyword")<>"" then 
	   url="kcgl.asp?keyword="&request("keyword")&"&type="&request("type")
	   sqlbody="SELECT * from xc where name like '%" & request("keyword") & "%' and wpid=0 and type="&request("type") 
       pagename=" 名称 "&request("keyword")
	end if 
	if request("class")<>"" then 
	   url="kcgl.asp?class="&request("class")&"&type="&request("type")
	   sqlbody="SELECT * from xc where  wpid=0 and zclass="&request("class")&" and type="&request("type")
       pagename="现存"
	end if 
	
	if request("sscj")<>"" then 
	   url="kcgl.asp?sscj="&request("sscj")&"&type="&request("type")
	   sqlbody="SELECT * from xc where sscj="&request("sscj")&" and type="&request("type")
       if request("sscj")=1000 then
	     pagename="分厂现存"
	   else
	     pagename="车间现存"
	   end if 	 
	end if 
	
	if request("action")="history" then 
	   url="kcgl.asp?action=history"&"&type="&request("type")
	   sqlbody="SELECT * from history where type="&request("type")&" and (wpid=0 or isnull(qx))"
	   pagename="历史记录"
	end if 

	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&title&"-"&pagename&"</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf
	call search()
	
	

	   sqlbody=sqlbody&" order by rcdate DESC"
	
	  set rsbody=server.createobject("adodb.recordset")
	  rsbody.open sqlbody,connkc,1,1
	  if rsbody.eof and rsbody.bof then 
		 message "<p align=""center"">暂无内容</p>" 
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
		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		 Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		 Dwt.out "<tr class=""x-grid-header"">"
		 Dwt.out "<td  class='x-td'><Div class='x-grid-hd-text'>编号</Div></td>"
		 Dwt.out "     <td  class='x-td' ><Div class='x-grid-hd-text'>车间</Div></td>"
		 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>分类</Div></td>"
		 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>名称</Div></td>"
		 Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>规格型号</Div></td>"
		 Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>单位</Div></td>"
		 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>单价</Div></td>"
		 Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>总数量</Div></td>"
Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>剩余数量</Div></td>"
		 Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>金额</Div></td>"
		 Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>更新时间</Div></td>"
		 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>操作人</Div></td>"
		 'Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>单据号</Div></td>"
		 'Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>存放地址</Div></td>"
		'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>订货</div></td>"
		'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>计划到货</div></td>"
		'dwt.out "      <td class='x-td'><DIV class='x-grid-hd-text'>实际到货</div></td>"
		 Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>备 注</Div></td>"
		 if request("action")<>"history" then Dwt.out "      <td class='x-td'><Div class='x-grid-hd-text'>选 项</Div></td>"
		 Dwt.out "    </tr>"
	  
	  do while not rsbody.eof and rowcount>0
			dim xh_id
				xh_id=((page-1)*pgsz)+1+xh
				xh=xh+1
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
			if request("action")<>"history" then  Dwt.out "<a href='#' onclick=""showsubmenu("&rsbody("id")&");"" id=xxx"&rsbody("id")&"><img src='/img_ext/i7.gif' /></a>"
			Dwt.out xh_id&"</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px""  ><Div align=""center"">"&sscjh(rsbody("sscj"))&"</Div></td>"
	
			Dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"" >"&dclass(rsbody("dclass"))&"-"&dclass(rsbody("zclass"))&"</td>"
			dim bjname
if rsbody("numb")>0 then 
bjname="<font color=#ff0000>"&rsbody("name")&"</font>"
else
bjname=rsbody("name")
end if 
Dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"" >"&bjname&"&nbsp;</td>"


			Dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"" >"&rsbody("xhgg")&"&nbsp;</td>"
			Dwt.out "  <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("dw")&"&nbsp;</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("dmoney")&"&nbsp;</Div></td>"
dim totalnumb,sqlk,rsk

	   sqlk="select numb from history where wpid="&rsbody("id")
	totalnumb=0
	  set rsk=server.createobject("adodb.recordset")
	  rsk.open sqlk,connkc,1,1
	  if rsk.eof and rsk.bof then 
		 totalnumb=rsbody("numb")
	  else

	   totalnumb=connkc.Execute(sqlk)(0)
	 end if
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&totalnumb&"&nbsp;</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("numb")&"&nbsp;</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&totalnumb*rsbody("dmoney")&"&nbsp;</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("rcdate")&"&nbsp;</Div></td>"
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&usernameh(rsbody("userid"))&"&nbsp;</Div></td>"
			'Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("crknumb")&"&nbsp;</Div></td>"
			'Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("adress")&"&nbsp;</Div></td>"
			
			'dhdate=rsbody("dhdate")
			'if dhdate="" or isnull(dhdate) then 
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未编辑&nbsp;</span></div></td>"
			'else
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate&"&nbsp;</div></td>"
			'end if
						
			'dhdate1=rsbody("jhdhdate")
			'if dhdate1="" or isnull(dhdate1) then 
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未编辑&nbsp;</span></div></td>"
			'else
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate1&"&nbsp;</div></td>"
			'end if

			dhdate2=rsbody("sjdhdate")
			'if dhdate2="" or isnull(dhdate2) then 
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center><span style=""color:#ff0000"">未编辑&nbsp;</span></div></td>"
			'else
			'dwt.out "      <td style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><div align=center>"&dhdate2&"&nbsp;</div></td>"
			'end if
			Dwt.out " <td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&rsbody("bz")&"&nbsp;</Div></td>"
		   if request("action")<>"history" then 
			   Dwt.out "<td style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"
			   call editdel(rsbody("id"),rsbody("sscj"),rsbody("numb"),rsbody("type"))
			   Dwt.out "</Div></td>"
		   end if 
		   Dwt.out "</tr>"
		   totalamoney=totalamoney+totalnumb*rsbody("dmoney")
            '金额的科学计数，每三位加一个逗号，未实现
'			if len(totalamoney)>4 then totalamoney=
	'查看历史记录时，不需要显示出库记录	
	if request("action")<>"history"  then 
		Dwt.out "<tr ><td  colspan=14 style='display:none' id='info"&rsbody("id")&"'>"		
		Dwt.out "<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		  dim rscj,sqlcj
		  sqlcj="SELECT * from history where wpid="&rsbody("id")&" order by rcdate DESC"
		  set rscj=server.createobject("adodb.recordset")
		  rscj.open sqlcj,connkc,1,1
		  if rscj.eof and rscj.bof then 
			 Dwt.out  "无出库记录" 
		  else
			Dwt.out "<tr >" & vbCrLf
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>出库对象</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>出库数量</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>操作人</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>单据号</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>时间</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>存放地址</Div></td>"
			Dwt.out "      <td  bgcolor='#BFDFFF'><Div class='x-grid-hd-text'>备注</Div></td>"
			Dwt.out  "    </tr>"
			  do while not rscj.eof
					dim xhcj
						xhcj=xhcj+1
						'on error resume next
					Dwt.out "<tr class='x-grid-row'  >"& vbCrLf
					Dwt.out "      <td  bgcolor='#BFDFFF'>"
					if rscj("qx")<>"" then 
					   if rscj("qx")="1000" then
					    Dwt.out "现场使用"
					   else
					    Dwt.out rscj("qx")
					   end if 
					else   	
					   Dwt.out sscjh(rscj("sscj"))
					end if 
					Dwt.out "</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&rscj("numb")&"&nbsp;</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&usernameh(rscj("userid"))&"&nbsp;</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&rscj("crknumb")&"&nbsp;</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&rscj("rcdate")&"&nbsp;</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&rscj("adress")&"&nbsp;</td>"
					Dwt.out "      <td  bgcolor='#BFDFFF'>"&rscj("bz")&"&nbsp;</td>"
					Dwt.out  "    </tr>"
				rscj.movenext
				loop
		end if 		
		Dwt.out "</table>"		
		Dwt.out "</tr>"		
   end if 
			RowCount=RowCount-1
		rsbody.movenext
		loop
			Dwt.out " <tr class='x-grid-row x-grid-row-alt' >"
			Dwt.out " <td ><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td ><Div align=""center""><font color=#FF0000>合计</font></Div></td>"
			Dwt.out "  <td>&nbsp;</td>"
			Dwt.out "  <td>&nbsp;</td>"
			Dwt.out "  <td >&nbsp;</td>"
			Dwt.out "  <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out "  <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out "  <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td ><Div align=""center""><font color=#FF0000>"&totalamoney&"</font>&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td><Div align=""center"">&nbsp;</Div></td>"
			Dwt.out " <td ><Div align=""center"">&nbsp;</Div></td>"
		   if request("action")<>"history"  then Dwt.out "<td><Div align=""center"">&nbsp;</Div></td></tr>"
	
	   Dwt.out "</table>"
	  
	  'if request("class")="" and request("sscj")="" then 
	'	 call showpage1(page,url,total,record,PgSz)
	  'else
		 call showpage(page,url,total,record,PgSz)
	  'end if 	 
		Dwt.out "</Div>"& vbCrLf
	end if
	Dwt.out "</Div>"  
	  rsbody.close
	  set rsbody=nothing
	  conn.close
	  set conn=nothing
end sub




Dwt.out "</body></html>"

'用于显示父分类名称 
Function dclass(classid)
	dim sqlname,rsname
	dim sqlz,rsz
	sqlz="SELECT * from class where id="&classid
    set rsz=server.createobject("adodb.recordset")
    rsz.open sqlz,connkc,1,1
    if rsz.eof then 
	  'Dwt.out "未分类"
	else  
	    dclass=rsz("name")
	end if 
	rsz.close
	set rsz=nothing
end Function 


'选项（编辑、出库\删除）
sub editdel(id,sscj,numb,cltype)
 'if session("levelclass")=sscj then 
	'Dwt.out "<a href=kcgl_fcsa.asp?action=sr&id="&id&">入库</a>&nbsp;"
	if displaypagelevelh(session("groupid"),2,session("pagelevelid")) then 
		if numb<>0 then Dwt.out "<a href=kcgl.asp?action=fc&id="&id&"&type="&cltype&">出库</a>&nbsp;"
	end if 
	if  displaypagelevelh(session("groupid"),3,session("pagelevelid")) then
		Dwt.out "<a href=kcgl.asp?action=edit&type="&request("type")&"&id="&id&" >编</a> "
		Dwt.out "<a href=kcgl.asp?action=del&type="&request("type")&"&id="&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
    end if 
' else
 '   Dwt.out "&nbsp;"
 'end if 
end sub



sub search()
    Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	'按名称搜索
	Dwt.out "  <form method='Get' name='form1' action='kcgl.asp'>" & vbCrLf
     
    Dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
		if request("keyword")<>"" then 
			Dwt.out "value='"&request("keyword")&"'"
			Dwt.out ">" & vbCrLf
		else
			Dwt.out "value='输入控索的名称'"
			Dwt.out " onblur=""if(this.value==''){this.value='输入控索的名称'}"" onfocus=""this.value=''"">" & vbCrLf
		end if    
		Dwt.out " <input name='type' type='hidden' value='"&request("type")&"'> <input type='submit' name='Submit'  value='搜索'>" & vbCrLf	
		Dwt.out "<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
	dim sqlcj,rscj
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
		set rscj=server.createobject("adodb.recordset")
		rscj.open sqlcj,conn,1,1
		do while not rscj.eof
			Dwt.out"<option value='kcgl.asp?type="&request("type")&"&sscj="&rscj("levelid")&"'"
			if cint(request("sscj"))=rscj("levelid") then Dwt.out" selected"
			Dwt.out">"&rscj("levelname")&"</option>"& vbCrLf	
			rscj.movenext
		loop
		Dwt.out"<option value='kcgl.asp?sscj=1000&type="&request("type")&"'>分厂</option>"& vbCrLf
		rscj.close
		set rscj=nothing
		Dwt.out "</select>"

	Dwt.out "   &nbsp;&nbsp;&nbsp;&nbsp;按分类显示：<select name='kcgl_dclass' size='1' id='cat1' onChange=""selectpc(this.value,'b',document.form1.kcgl_zclass)"">"& vbCrLf
	Dwt.out "  <option selected value='0'>选择一级分类</option>"& vbCrLf
	sql="SELECT * from class where dclass=0 and type="&request("type")
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
    do while not rs.eof
       	Dwt.out"<option value='"&rs("id")&"'>"&rs("name")&"</option>"& vbCrLf
		rs.movenext
	loop
	rs.close
	set rs=nothing
	Dwt.out "</select>"& vbCrLf
	Dwt.out "<select name='kcgl_zclass' size='1' id='cat2'  onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">"& vbCrLf
	Dwt.out "  <option selected value='0'>选择二级分类</option>"& vbCrLf
	Dwt.out "</select>"& vbCrLf
	Dwt.out "<script language='javascript'>"& vbCrLf
	Dwt.out "function selectpc(parentValue,child,addObj){"& vbCrLf


    dim b,bv,b_p,sqlz,rsz
	sql="SELECT * from class where dclass=0 "& vbCrLf
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connkc,1,1
         b="var b =   new Array("
        bv="var bv =   new Array("
        b_p="var b_p =   new Array("
   
	do while not rs.eof
		sqlz="SELECT * from class where dclass="&rs("id")&" and type="&request("type")&" order by orderby"
        set rsz=server.createobject("adodb.recordset")
        rsz.open sqlz,connkc,1,1
        if rsz.eof and rsz.bof then
		   b=b&"'无二级分类',"
		   bv=bv&"'kcgl.asp?type="&request("type")&"',"
		   b_p=b_p&"'"&rs("id")&"',"
		else
		do while not rsz.eof
			
			b=b&"'"&rsz("name")&"',"
			bv=bv&"'kcgl.asp?type="&request("type")&"&class="&rsz("id")&"',"
			b_p=b_p&"'"&rs("id")&"',"
		   rsz.movenext
	    loop
	    end if 
		rsz.close
	    set rsz=nothing
		rs.movenext
	loop
	rs.close
	set rs=nothing
	b=left(b,len(b)-1)
	bv=left(bv,len(bv)-1)
	b_p=left(b_p,len(b_p)-1)
	b=b&");"
	bv=bv&");"
	b_p=b_p&");"
	Dwt.out b & vbCrLf
	Dwt.out bv & vbCrLf
	Dwt.out b_p & vbCrLf
	
	
	
	Dwt.out "var labelValue = new Array();"& vbCrLf
	Dwt.out "var labelText =  new Array();"& vbCrLf
	Dwt.out "var k = 0;"& vbCrLf
	
	Dwt.out "cObj = eval(child);"& vbCrLf
	Dwt.out "cObjV = eval(child+'v');"& vbCrLf
	Dwt.out "cpObj = eval(child + '_p');"& vbCrLf
	Dwt.out "for(i=0; i<cpObj.length; i++)"& vbCrLf
	Dwt.out "{"& vbCrLf
	Dwt.out "	if(cpObj[i] == parentValue)"& vbCrLf
	Dwt.out "	{"& vbCrLf
	Dwt.out "		labelText[k] =  cObj[i];"& vbCrLf
	Dwt.out "		labelValue[k] =	cObjV[i]; "& vbCrLf
	Dwt.out "		k++;"& vbCrLf
	Dwt.out "	}"& vbCrLf
	Dwt.out "}"& vbCrLf
	
	
	Dwt.out "addObj.options.length = 0;"& vbCrLf
	Dwt.out "addObj.options[0] = new Option('选择二级分类','0');"& vbCrLf
	Dwt.out "for(i = 0; i < labelText.length; i++) {"& vbCrLf
	Dwt.out "	addObj.add(document.createElement('option'));"& vbCrLf
	Dwt.out "	addObj.options[i+1].text=i+1+'  '+labelText[i];"& vbCrLf
	Dwt.out "	addObj.options[i+1].value=labelValue[i];"& vbCrLf
	Dwt.out "}"& vbCrLf
	Dwt.out "addObj.selectedIndex = 0;"& vbCrLf
    Dwt.out "}"& vbCrLf
    Dwt.out "</script>"& vbCrLf

	Dwt.out "</form></Div></Div>"

end sub


Call CloseConn
%>