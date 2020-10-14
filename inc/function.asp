<!--#include file="conn.asp"-->
<%
Dim dwt
Set dwt= New dwt_Class	
Class dwt_Class
	Public Function out(s) 
		response.write s
	End Function 


    '保存系统记录
	Public Function savesl(leftname,action,message)
		dim leftmdb,connleft,connl
		leftmdb="ybdata/left.mdb"
		Set connleft = Server.CreateObject("ADODB.Connection")
		connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
		connleft.Open connl    

	  dim rsadd,sqladd,trueip
	  set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from systemlog" 
      rsadd.open sqladd,connleft,1,3
      rsadd.addnew
	  TrueIP = Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
      If TrueIP = "" Then TrueIP = Request.ServerVariables("REMOTE_ADDR")
	  rsadd("ip")=trueip
	  rsadd("userid")=session("userid")
      rsadd("message")=message
      rsadd("action")=action
      rsadd("leftname")=leftname
      rsadd("update")=now()
	  rsadd.update
      rsadd.close
      set rsadd=nothing

	end Function
	
	
	Public Function pagetop(title) 
		dwt.out "<html>"& vbCrLf
		dwt.out "<head>" & vbCrLf
		dwt.out "<title>"&title&"</title>"& vbCrLf
		dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
		dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
		dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
		dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"
		dwt.out "</head>"& vbCrLf
		dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf
    end function
   
   
   
    'formname:表单名称
	'inputname:要比较的表单的名称
	'alerttext:弹出界面的文字
	'checkvalue:比较的值
	Public Function formcheck(formname,inputname,alerttext,checkvalue) 
		dwt.out "  if(document."&formname&"."&inputname&".value=="&checkvalue&"){" & vbCrLf
		dwt.out "      alert('"&alerttext&"');" & vbCrLf
		dwt.out "  document.form1."&inputname&".focus();" & vbCrLf
		dwt.out "      return false;" & vbCrLf
		dwt.out "    }" & vbCrLf
	end function	
	
	
	
	
	'输出表单头
	'url:action的地址
	'forname：表单名称
	'title:表单标题 
	'checkname:检查表单内容的名称
	Public Function lable_title(url,formname,title,checkname) 
		dwt.out"<DIV style='WIDTH: 760px;padding-top:50px;padding-left:100px'>"& vbCrLf
		dwt.out"  <DIV class=x-box-tl>"& vbCrLf
		dwt.out"	<DIV class=x-box-tr>"& vbCrLf
		dwt.out"	  <DIV class=x-box-tc></DIV>"& vbCrLf
		dwt.out"	</DIV>"& vbCrLf
		dwt.out"  </DIV>"& vbCrLf
		dwt.out"  <DIV class=x-box-ml>"& vbCrLf
		dwt.out"	<DIV class=x-box-mr>"& vbCrLf
		dwt.out"	  <DIV class=x-box-mc>"& vbCrLf
		dwt.out"		<div align=center><H3 style='MARGIN-BOTTOM: 5px'>"&title&"</H3></div>"& vbCrLf
		dwt.out"		<DIV id=form-ct>"& vbCrLf
		dwt.out "<form method='post' action='"&url&"' name='"&formname&"' onsubmit='javascript:return "&checkname&"();'>"
		dwt.out"			<DIV class='x-form-ct'>"& vbCrLf
	End Function 
    
	'输出INPUT,
	'leftname:input在页面上显示的名称
	'inputname:input在表单中的名称
	'inputformvalue:input的值,在表单中传递用
	'inputvaluename:input的值,在页面中显示
	'isdisabled:input是否为禁用
	'isbt:是否必添项
	'tips:提示信息
	Public Function lable_input(leftname,inputname,inputformvalue,inputvaluename,isdisabled,isbt,tips)
		dwt.out"<DIV class='x-form-item'>"& vbCrLf
		dwt.out"<LABEL style='WIDTH: 75px'><div align=right>"&leftname&":</div></LABEL>"& vbCrLf
		dwt.out"<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
		dwt.out"<input size=20 name='"&inputname&"' type='text' value='"&inputvaluename&"'"
		if isdisabled then  dwt.out "disabled='disabled'"
		dwt.out " >"
		if isdisabled then dwt.out "<input name='"&inputname&"' type='hidden' value='"&inputformvalue&"'>"
		if isbt then dwt.out " <span class='red'>*</span>"
		dwt.out " <span class='tips'>"&tips&"</span>"
		dwt.out"</DIV>"& vbCrLf
		dwt.out"</DIV>"& vbCrLf
		dwt.out"<DIV class=x-form-clear-left></DIV>"& vbCrLf
    end function

	'输出自动完成需要的头文件
	Public Function complete_a 
		dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
		dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
		dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	End Function 

	'输出带自动完成的INPUT,
	'leftname:input在页面上显示的名称
	'inputname:input在表单中的名称
	'isbt:是否必添
	'tips:input后的提示信息
	'--------------------------------
	'dbname:要自动完成的数据库名称
	'zdtext:字段名
	'btext:表名称
	Public Function lable_input_complete(leftname,inputname,isbt,tips,dbname,zdtext,btext)
		dwt.out"<DIV class='x-form-item'>"& vbCrLf
		dwt.out"<LABEL style='WIDTH: 75px;'><div align=right>"&leftname&":</div></LABEL>"& vbCrLf
		dwt.out"<DIV class='x-form-element'  style='PADDING-LEFT: 80px'>"& vbCrLf
		dwt.out"<input name='"&inputname&"' >"
		if isbt then dwt.out " <span class='red'>*</span>"
		dwt.out " <span class='tips'>"&tips&"</span>"
		dwt.out"</DIV>"& vbCrLf
		dwt.out"</DIV>"& vbCrLf
		dwt.out"<DIV class=x-form-clear-left></DIV>"& vbCrLf

		
		dwt.out "<script>"
		'自动完成后面的内容为选中状态
		dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
		dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
		dwt.out "new CAPXOUS.AutoComplete("""&inputname&""", function() {return ""/inc/autocomplete.asp?dbname="&dbname&"&zdtext="&zdtext&"&btext="&btext&"&typing="" + escape(this.text.value);});"
		dwt.out "</script>"
    end function

	'输出INPUT,时间框
	'leftname:input在页面上显示的名称
	'inputname:input在表单中的名称
	'inputvalue:时间的值
	'isbt:是否必添
	'tips:input后的提示信息
	'--------------------------------
	Public Function lable_input_date(leftname,inputname,inputvalue,isdisabled,isbt,tips)
		dwt.out"			  <DIV class='x-form-item'>"& vbCrLf
		dwt.out"				<LABEL style='WIDTH: 75px'><div align=right>"&leftname&":</div></LABEL>"& vbCrLf
		dwt.out"				<DIV class='x-form-element' style='PADDING-LEFT: 80px'>"& vbCrLf
		dwt.out"<input name='"&inputname&"' type='text'  onClick='new Calendar(0).show(this)' readOnly  value='"&inputvalue&"'"
		if isdisabled then  dwt.out "disabled='disabled'"
		dwt.out " >"
		if isbt then dwt.out " <span class='red'>*</span>"
		dwt.out" <span class='tips'>"&tips&"</span>"& vbCrLf
		dwt.out"				</DIV>"& vbCrLf
		dwt.out"			  </DIV>"& vbCrLf
		dwt.out"			  <DIV class=x-form-clear-left></DIV>"& vbCrLf
    end function
	
	
	
	
	
	'输出表单尾
	'action:action的名称
	'submitname:按钮的名称
	'isid:是否带有ID参数，用于编辑修改
	'idname:还的ID的NAME
	'ID:标识ID
	Public Function lable_footer(action,submitname,isid,idname,id) 
		dwt.out" <DIV class=x-form-clear></DIV>"& vbCrLf
		dwt.out"			</DIV>"& vbCrLf
		dwt.out"			<DIV class=x-form-btns-ct>"& vbCrLf
		dwt.out"			  <DIV class='x-form-btns x-form-btns-center'><div align=center>"& vbCrLf
		if isid then  dwt.out "<input type='hidden' name='"&idname&"' value='"&id&"'>  "
		dwt.out"<input name='action' type='hidden' id='action' value='"&action&"'><input  type='submit' name='Submit' value='"&submitname&"' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
		dwt.out"				<DIV class=x-clear></DIV>"& vbCrLf
		dwt.out"			  </div></DIV>"& vbCrLf
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
	End Function 
    
	Public function search_sscj(url,isfc)
		dim sqlcj,rscj
		dwt.out "<select name='JumpClass' id='JumpClass' onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
		dwt.out "	       <option value=''>按车间跳转至…</option>" & vbCrLf
		sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
			set rscj=server.createobject("adodb.recordset")
			rscj.open sqlcj,conn,1,1
			do while not rscj.eof
				dwt.out"<option value='"&url&rscj("levelid")&"'"
				if cint(request("sscj"))=rscj("levelid") then dwt.out" selected"
				dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf
				rscj.movenext
			loop
				if isfc then dwt.out"<option value='kcgl.asp?sscj=1000'>分厂</option>"& vbCrLf
			rscj.close
			set rscj=nothing
			dwt.out "     </select>	" & vbCrLf
	End Function 


	Public function search_key(inputvalue)
		dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
		if request("keyword")<>"" then 
			dwt.out "value='"&request("keyword")&"'"
			dwt.out ">" & vbCrLf
		else
			dwt.out "value='"&inputvalue&"'"
			dwt.out " onblur=""if(this.value==''){this.value='"&inputvalue&"'}"" onfocus=""this.value=''"">" & vbCrLf
		end if    
		dwt.out "  <input type='submit' name='Submit'  value='搜索'>" & vbCrLf
    end function
End Class







'↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓120711正则提取位号并识别,哪果有检修更换周检日志则替换 

'  html值班日志内容 
'sscj所属车间zblogdate值班日志日期,这两个参数在此模块不用，在下一个功能块用
Public Function whsb(html,sscj,zblogdate)
	'html=uCase(html)  '所有英文转为大写
	'html1=RegExpTest("(-?\d*)(\.\d+)?", html)  '所有数字
	
'html1=RegExpTest("\d{5}|\d{3}", html)  '3到5位数字
'html1=RegExpTest("([a-zA-z]{1,3})(-|_)(\d{2,5}\w{1,3})", html)  '3到5位数字



Dim regEx, Match, Matches ' 建立变量。
  Set regEx = New RegExp ' 建立正则表达式。
  regEx.Pattern = "([a-zA-z]{1,3})(-|_)(\d{2,5}\w{1,3})" ' 位号识别的正则。
  regEx.IgnoreCase = false ' 设置是否区分字符大小写。
  regEx.Global = True ' 设置全局可用性。
  
 'RetStr=regex.Replace(strng,"<b>$1$2$3</b>") '替换
  
  '提取
  Set Matches = regEx.Execute(html) ' 执行搜索。
  For Each Match in Matches ' 遍历匹配集合。
  'RetStr = RetStr & "Match found at position "
 ' RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
	if match.value<>""  then RetStr = RetStr & Match.Value&","   '列出所有符合条件的值
  'if match.value<>"" then RetStr =  replace(strng,match.value,"<b>sdfsdfsdf</b>")&"<br>"&match.value
  
  Next
  	
  RegExpTest = replacewh(html,sscj,zblogdate,RetStr)
  		Set regex = Nothing	



	whsb = RegExpTest
End Function






'查找搜索出来的位号的检修更换周检，如果有记录则替换
'HTML2值班日志内容 sscj所属车间 zblogdate日期，filter要替换的位号
Public Function replacewh(html2,sscj,zblogdate, filter)
	'html=LCase(html)
	filter=split(filter,",")
	For Each iiii In filter
		''''''调试代码
		'html3=html3&iiii&" "
  ' i=i+1
	'response.write i&"<br>"
		''''''调试代码
	
	
	'提取位号中的数字
	Dim regEx11' 建立变量。
  Set regEx11 = New RegExp ' 建立正则表达式。
  regEx11.Pattern = "(\d{5}|\d{3})" ' 3到5位数字。
  regEx11.IgnoreCase = false ' 设置是否区分字符大小写。
  regEx11.Global = True ' 设置全局可用性。
  '提取
  Set Matches = regEx11.Execute(iiii) ' 执行搜索。
  For Each Match in Matches ' 遍历匹配集合。
		'whnumb=whnumb&iiii&" "&Match.Value&"<br>" '测试用 在位号中找到 的3到5位数字
		  whnumb=Match.Value    '位号中的数字
		 
		 
		 
		 
		 
		 '根据位号中的数字，车间，检修日期 查询检修记录中的个数
		  sqljx = "SELECT sbjx.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id"
		  sqljx=sqljx&"  WHERE (((sb.sb_sscj)="&sscj&")) and (((sb.sb_wh) like '%" &whnumb& "%')) "
			sqljx=sqljx&" and jx_date=#"&zblogdate&"#"	
			'sqljx=sqljx&" order by  jx_date DESC"	
			'dwt.out sqljx&"<br>"
			  set rsjx=server.createobject("adodb.recordset")
			  rsjx.open sqljx,conn,1,1
			  if rsjx.eof and rsjx.bof then 
				  record_jx=0
			  else
				  record_jx=rsjx.recordcount  '查到的总数
				  
				  
			  end if
			  rsjx.close
			  set rsjx=nothing
	   
	   
	   sqljx = "SELECT sbgh.*, sb.sb_sscj,sb.sb_wh,sb.sb_dclass FROM sbgh INNER JOIN sb ON sbgh.sb_id = sb.sb_id "
		  sqljx=sqljx&"  WHERE (((sb.sb_sscj)="&sscj&")) and (((sb.sb_wh) like '%" &whnumb& "%')) "
			sqljx=sqljx&" and gh_date=#"&zblogdate&"#"	
			'sqljx=sqljx&" order by  jx_date DESC"	
			'dwt.out sqljx&"<br>"
			  set rsjx=server.createobject("adodb.recordset")
			  rsjx.open sqljx,conn,1,1
			  if rsjx.eof and rsjx.bof then 
				  record_gh=0
			  else
				  record_gh=rsjx.recordcount  '查到的总数
			  end if
			  rsjx.close
			  set rsjx=nothing
	 
	 
	 
	 '查询日期时间有问题，周检暂不做了  
'	   sqljx = "SELECT zjinfo.*, zjtz.sscj,zjtz.wh FROM zjinfo INNER JOIN zjtz ON zjinfo.zjtzid = zjtz.id "
'		  sqljx=sqljx&"  WHERE (((zjtz.sscj)="&sscj&")) and (((zjtz.wh) like '%" &whnumb& "%')) "
'			sqljx=sqljx&" and zjinfo.zjdate=#2011-05-14#"	
'			'sqljx=sqljx&" order by  jx_date DESC"	
'			'dwt.out sqljx&"<br>"
'			  set rsjx=server.createobject("adodb.recordset")
'			  rsjx.open sqljx,connzj,1,1
'			  if rsjx.eof and rsjx.bof then 
'				  record_zj=0
'			  else
'				  record_zj=rsjx.recordcount  '查到的总数
'			  end if
'			  rsjx.close
'			  set rsjx=nothing
'	   response.Write sqljx&"<br>"
	    'if record_jx<>0 or record_gh<>0 or record_zj<>0 then html2=replace(html2,iiii,"<b>"&iiii&"(检"&record_jx&",换"&record_gh&",周"&record_zj&")</b>")   '看是否有检修替换原始内容中的位号
	   if record_jx<>0 or record_gh<>0 then html2=replace(html2,iiii,"<b>"&iiii&"(检"&record_jx&",换"&record_gh&")</b>")   '看是否有检修替换原始内容中的位号
	   
		 ' html2=replace(html2,iiii,"<b>"&iiii&" "&sscj&" "&zblogdate&"</b>")   '测试用替换原始内容中的位号
 Next
	
	
	

	
	
	
	Next
		Set regex11 = Nothing	
	replacewh = html2
End Function



'↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑120711正则提取位号并识别,哪果有检修更换周检日志则替换 







Dim Action, FoundErr, ErrMsg, ComeUrl,total1
Dim strInstallDir


'****************************
'文件结尾版权
'***********************888
sub footer()
response.write "<br><br>"
response.write "<table cellpadding=""2"" cellspacing=""1"" border=""0"" width=""100%"" class=""border"" align=center>"
response.write "<tr align=""center"">"
response.write "<td height=25 class=""topbg""><span class=""Glow"">设备管理系统 All Rights Reserved.</span>"
response.write "</tr></table></body></html>"
end sub



'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Function ReplaceBadChar(strChar)
'    strChar=REPLACE(STRCHAR,"'","")
'    ReplaceBadChar = strChar
'131123加
 If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
   Dim strBadChar, arrBadChar, tempChar, I
   'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,,;,:," & Chr(34) & "," & Chr(0) & ""
   strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,,;,:," & Chr(34) & "," & Chr(0) & ""
   arrBadChar = Split(strBadChar, ",")
   tempChar = strChar
   For I = 0 To UBound(arrBadChar)
    tempChar = Replace(tempChar, arrBadChar(I), "")
   Next
   tempChar = Replace(tempChar, "@@", "@")
   ReplaceBadChar = tempChar	
End Function

Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Function PE_CDbl(ByVal str1)
    If IsNumeric(str1) Then
        PE_CDbl = CDbl(str1)
    Else
        PE_CDbl = 0
    End If
End Function



'专门用来去除内容中的文本害码。。。
Public Function DecodeFilter(html, filter)
	'html=LCase(html)
	filter=split(filter,",")
	For Each iiii In filter
		Select Case iiii
			Case "SCRIPT"		' 去除所有客户端脚本javascipt,vbscript,jscript,js,vbs,event,...
				html = exeRE("(javascript|jscript|vbscript|vbs):", "#", html)
				html = exeRE("</?script[^>]*>", "", html)
				html = exeRE("on(mouse|exit|error|click|key)", "", html)
			Case "TABLE":		' 去除表格<table><tr><td><th>
				html = exeRE("</?table[^>]*>", "", html)
				html = exeRE("</?tr[^>]*>", "", html)
				html = exeRE("</?th[^>]*>", "", html)
				html = exeRE("</?td[^>]*>", "", html)
				html = exeRE("</?tbody[^>]*>", "", html)
			Case "CLASS"		' 去除样式类class=""
				html = exeRE("(<[^>]+) class=[^ |^>]*([^>]*>)", "$1 $2", html) 
			Case "STYLE"		' 去除样式style=""
				html = exeRE("(<[^>]+) style=""[^""]*""([^>]*>)", "$1 $2", html)
				html = exeRE("(<[^>]+) style='[^']*'([^>]*>)", "$1 $2", html)
			Case "IMG"		' 去除样式style=""
				html = exeRE("</?img[^>]*>", "", html)
			Case "XML"		' 去除XML<?xml>
				html = exeRE("<\\?xml[^>]*>", "", html)
			Case "NAMESPACE"	' 去除命名空间<o:p></o:p>
				html = exeRE("<\/?[a-z]+:[^>]*>", "", html)
			Case "FONT"		' 去除字体<font></font>
				html = exeRE("</?font[^>]*>", "", html)
			Case "A"		' 去除字体<font></font>
				html = exeRE("</?a[^>]*>", "", html)
			Case "MARQUEE"		' 去除字幕<marquee></marquee>
				html = exeRE("</?marquee[^>]*>", "", html)
			Case "OBJECT"		' 去除对象<object><param><embed></object>
				html = exeRE("</?object[^>]*>", "", html)
				html = exeRE("</?param[^>]*>", "", html)
				'html = exeRE("</?embed[^>]*>", "", html)
			Case "EMBED"
			   html =  exeRE("</?embed[^>]*>", "", html)
			Case "DIV"		' 去除对象<object><param><embed></object>
				html = exeRE("</?div([^>])*>", "$1", html)
			Case "STRONG"		' 去除对象<object><param><embed></object>
				html = exeRE("</?strong([^>])*>", "$1", html)
			Case "ONLOAD"		' 去除样式style=""
				html = exeRE("(<[^>]+) onload=""[^""]*""([^>]*>)", "$1 $2", html)
				html = exeRE("(<[^>]+) onload='[^']*'([^>]*>)", "$1 $2", html)
			Case "ONCLICK"		' 去除样式style=""
				html = exeRE("(<[^>]+) onclick=""[^""]*""([^>]*>)", "$1 $2", html)
				html = exeRE("(<[^>]+) onclick='[^']*'([^>]*>)", "$1 $2", html)
			Case "ONDBCLICK"		' 去除样式style=""
				html = exeRE("(<[^>]+) ondbclick=""[^""]*""([^>]*>)", "$1 $2", html)
				html = exeRE("(<[^>]+) ondbclick='[^']*'([^>]*>)", "$1 $2", html)
			
	
		End Select
	Next
	'html = Replace(html,"<table","<")
	'html = Replace(html,"<tr","<")
	'html = Replace(html,"<td","<")
	DecodeFilter = html
End Function

'正则替换。。。
Public Function exeRE(re, rp, content)
	Set oReg = New RegExp
	oReg.IgnoreCase =True
	oReg.Global=True	
	oReg.Pattern=re
	r = oReg.Replace(content,rp)
	Set oReg = Nothing	
	exeRE = r
End Function




'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
'URL中带？的
'*******************************************
sub showpage(page,url,total,record,pgsz)
   response.write "<div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  if Total =1 then 
    response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  else
   response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div>"
end sub




'********************************************8
'分页显示page当前页数，url网页地址，total总页数 record总条目数
'pgsz 每页显示条目数
' url中不带？
'*******************************************
sub showpage1(page,url,total,record,pgsz)
   response.write "<div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"?page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"?pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"?pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"?pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
   if Total =1 then 
     response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">条"
   else
     response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"?pgsz='+this.value;"">条"
   end if 
   if Total=1 then 
       response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
       response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"?pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div>"
end sub



sub showpage2(page,url,total,record,pgsz)
   response.write "<div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1><img src=../images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=../images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=../images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=../images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  if Total =1 then 
    response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  else
   response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">条"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div>"
end sub



'1维修一车间，2维修二车间，3维修三车间，4维修四车间，5综合车间，6计量科
Function sscjh(sscj)
    dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    if rscj.eof then 
	  sscjh="未知"
	else  
	do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    sscjh=rscj("levelname")
		rscj.movenext
	loop
	end if
	rscj.close
	set rscj=nothing
	if sscj=1000 then sscjh=" 分厂"
end Function 

'用于短的车间显示
Function sscjh_d(sscj)
       dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&sscj
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	'response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	    sscjh_d=replace(replace(replace(rscj("levelname"),"修",""),"车间",""),"科","")
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
end Function 

 '29日新加
'用于编辑新增装置显示
function formgh(ghid,sscj)
	dim sqlgh,rsgh
	
'
	if isnull(sscj) then sscj=0
	if isnull(ghid) then ghid=0
	if sscj=4 or sscj=5 then 
		sqlgh="SELECT * from ghname"
	else
	sqlgh="SELECT * from ghname where sscj="&sscj
	end if 		
	set rsgh=server.createobject("adodb.recordset")
			rsgh.open sqlgh,conn,1,1
			if rsgh.eof then 
			formgh="未编辑"
		else
			response.write"<option value='0'"
			if ghid=0 then response.write " selected" 
			response.write">请选择装置</option>"
			do while not rsgh.eof
				response.write"<option value='"&rsgh("ghid")&"' "
				if ghid=rsgh("ghid") then response.write "selected"
				response.write">"&rsgh("gh_name")&"</option>"  & vbCrLf   
			rsgh.movenext
		loop
	end if 
		rsgh.close
		set rsgh=nothing

end function
 '29日新加
'取装置工号名称
Function gh(ghid)
       dim sqlgh,rsgh
	if isnull(ghid) then ghid=0
	sqlgh="SELECT * from ghname where ghid="&ghid
    set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    if rsgh.eof then 
	  gh="未编辑"
	else
	    gh=rsgh("gh_name")
end if 
	rsgh.close
	set rsgh=nothing
end Function 
 '29日新加
'取分级的星数
Function fj(fjnumb)
       dim fj_i
	if isnull(fjnumb) or fjnumb=0 then 
	  fj="未分级"
	else
		for fj_i=1 to fjnumb
		fj=fj&"★"
		next
'	  if fjnumb=1 then fj="***"
'	  if fjnumb=2 then fj="**"
'	  if fjnumb=3 then fj="*"
	end if 
end Function 





'热动班1,供水班2,合成一班3,合成二班4,气压班5,复肥班6,硝铵班7,硝酸班8
Function ssbzh(ssbz)
            dim sqlbz,rsbz
	  sqlbz="SELECT * from bzname where id="&ssbz
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	else
	do while not rsbz.eof
       	'response.write"<option value='"&rsbz("levelid")&"'>"&rsbz("levelname")&"</option>"& vbCrLf
	    ssbzh=rsbz("bzname")
		rsbz.movenext
	loop
	end if 
	rsbz.close
	set rsbz=nothing
end Function
Function usernameh(userid)
       dim sqlbz,rsbz
	if isnull(userid) then userid=0
	sqlbz="SELECT * from userid where id="&userid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	else
	do while not rsbz.eof
       	'response.write"<option value='"&rsbz("levelid")&"'>"&rsbz("levelname")&"</option>"& vbCrLf
	    usernameh=rsbz("username1")
		rsbz.movenext
	loop
	end if 
	rsbz.close
	set rsbz=nothing
end Function
Function useridh(userid)
       dim sqlbz,rsbz
	if isnull(userid) then userid=0
	sqlbz="SELECT * from userid where id="&userid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	else
	do while not rsbz.eof
       	'response.write"<option value='"&rsbz("levelid")&"'>"&rsbz("levelname")&"</option>"& vbCrLf
	    useridh=rsbz("username")
		rsbz.movenext
	loop
	end if 
	rsbz.close
	set rsbz=nothing
end Function

''''''''**********************此两功能在usermanagement.asp页面中使用********************************************************88
'函数名称：checkpagelevelh 页面权限是否选择
'作用：usermanagement.asp判断用户在此页面的相关权限，在此页面有相关权限，则输出checked
'参数：userid用户ID，cz操作的参数（1查看2新建3编辑4删除），pageid页面ID
Function checkpagelevelh(groupid,cz,pageid)
    dim check_new,check_page1,i
	 check_new=split(pagelevelh(groupid),"/")
	 check_page1=split(check_new(cz),",")
	 For i = LBound(check_page1) To UBound(check_page1)
		if cint(check_page1(i))=cint(pageid) then dwt.out "checked"
	 Next 
end Function
'函数名称：pagelevelh
'作用：读出userid的pagelevel字段的对应值
'参数：userid用户ID
Function pagelevelh(groupid)
    dim sqlbz,rsbz
	sqlbz="SELECT * from grouplevel where id="&groupid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	else
	    pagelevelh=rsbz("pagelevel")
	end if 
	rsbz.close
	set rsbz=nothing
end Function
''''''''**********************************此两功能在usermanagement.asp页面中使用********************************************88


''''''''**********************此两功能在usermanagement.asp页面中使用********************************************************88
'函数名称：checkgrouplevelh 页面权限是否选择
'作用：usermanagement.asp判断用户在此页面的相关权限，在此页面有相关权限，则输出checked
'参数：groupid权限组，cz操作的参数（1查看2新建3编辑4删除），levelid是否有所属权限
Function checkgrouplevelh(groupid,cz,levelid)
	dim check_new,check_group1,i
	 check_new=split(grouplevelh(groupid),"/")
	 check_group1=split(check_new(cz),",")
	 For i = LBound(check_group1) To UBound(check_group1)
		if cint(check_group1(i))=cint(levelid) then dwt.out " checked "
		'response.Write(check_new(cz)&"fdsdfsdfs"&cint(groupid)&"fdsdfsdfs"&cint(check_group1(i)))
	 Next 
end Function
'函数名称：grouplevelh
'作用：读出groupid的grouplevel字段的对应值
'参数：groupid组ID
Function grouplevelh(groupid)
    dim sqlbz,rsbz
	sqlbz="SELECT * from grouplevel where id="&groupid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	else
	    grouplevelh=rsbz("grouplevel")
	end if 
	rsbz.close
	set rsbz=nothing
end Function


Function newsclassh(classid)
	sqlbz="SELECT * from xzgl_news_class where id="&classid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,connxzgl,1,1
    if rsbz.eof then 
	else
	    newsclassh=rsbz("class_name")
	end if 
	rsbz.close
	set rsbz=nothing
end Function
''''''''**********************************此两功能在usermanagement.asp页面中使用********************************************88







'*********************************8页面权限***********************************8
'函数名称：truepagelevelh
'作用：所有页面中判断用户在此页面的相关权限，如有输出TRUE，没有则弹出提示框
'参数：userid用户ID
Function truepagelevelh(groupid,cz,pageid)
    if pageid="" then 
	 message "无权访问此页面"
	else 
	 dim check_new,check_page1,i,pageleveltext
     truepagelevelh=false
	 pageleveltext=conn.Execute("SELECT pagelevel FROM grouplevel WHERE id="&groupid)(0)

	 check_new=split(pageleveltext,"/")
	 check_page1=split(check_new(cz),",")
	 For i = LBound(check_page1) To UBound(check_page1)
		if isnull(check_page1(i))=false then 
		 if cint(check_page1(i))=cint(pageid) then truepagelevelh=true
		end if 
     Next
	 		 'message session("pageleveltext")&"<br>"
	if truepagelevelh=false then  message "无权访问此页面"
   end if 	
end Function

'函数名称：displaypagelevelh
'作用：所有页面中判断用户在此页面的相关权限，如有输出TRUE，没有则输出false
'参数：userid用户ID
Function displaypagelevelh(groupid,cz,pageid)
dim check_new,check_page1,i,pageleveltext,rspage,sqlpage
'pageleveltext=conn.Execute("SELECT pagelevel FROM grouplevel WHERE id="&groupid)(0)
	
	
	'得到所在分组的页面权限
	Set rspage = Server.CreateObject("adodb.recordset")
    sqlpage = "select pagelevel from grouplevel where id="&groupid
    rspage.Open sqlpage, Conn, 1, 3
    If rspage.bof And rspage.EOF Then
          'response.write"<Script Language=Javascript>window.alert('用户名或密码错误!');location.href='index.asp';"
         Exit function
	else
	     pageleveltext=rspage("pagelevel")
    End If

'得到用用户所属车间，以判断在页面中是否有该车间的编辑权限，是否要用代定
'grouplevelid=conn.Execute("SELECT levelid FROM userid WHERE id="&session("userid")

displaypagelevelh=false
	 check_new=split(pageleveltext,"/")
	 check_page1=split(check_new(cz),",")
	 For i = LBound(check_page1) To UBound(check_page1)
		if isnull(check_page1(i))=false then 
		 if cint(check_page1(i))=cint(pageid) then displaypagelevelh=true
         '上句多加条件可一并判断页面权限，是否要用代定		
		end if 
	 Next 
end Function
'*********************************8页面权限***********************************8

'函数名称：displaygrouplevelh，组权限判断
'作用：所有页面中判断用户在此页面的"编辑删除"是否显示，如有输出TRUE，没有则输出false
'参数：groupid组ID，levelid所属车间ID
Function displaygrouplevelh(groupid,cz,levelid)
dim check_new,check_group1,i,groupleveltext
displaygrouplevelh=false
groupleveltext=conn.Execute("SELECT grouplevel FROM grouplevel WHERE id="&groupid)(0)
	 check_new=split(groupleveltext,"/")
	 check_group1=split(check_new(cz),",")
	 For i = LBound(check_group1) To UBound(check_group1)
		if isnull(check_group1(i))=false then 
		 if cint(check_group1(i))=cint(levelid) then displaygrouplevelh=true
		  'message cint(check_group1(i))&"gggg"&cint(groupid)
		end if 
	 Next 
end Function









'选项（编辑、删除）
sub editdel(id,sscj,editurl,delurl)
'displaypagelevelh页面权限和displaygrouplevelh组权限必须同时具备 
 if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,sscj) then 
    dwt.out "<a href="&editurl&id&">编辑</a>&nbsp;"
 end if 	
 if displaypagelevelh(session("groupid"),3,session("pagelevelid")) and displaygrouplevelh(session("groupid"),1,sscj) then 
     dwt.out "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 end if 
 dwt.out "&nbsp;"
end sub

'选项（编、删）
sub editdel_d(id,sscj,editurl,delurl)
'displaypagelevelh页面权限和displaygrouplevelh组权限必须同时具备 
 if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,sscj) then 
    dwt.out "<a href="&editurl&id&">编辑</a>&nbsp;"
 end if 	
 if displaypagelevelh(session("groupid"),3,session("pagelevelid")) and displaygrouplevelh(session("groupid"),1,sscj) then 
     dwt.out "<a href="&delurl&id&" onClick=""return confirm('确定要删除此记录吗？');"">删除</a>"
 end if 
 dwt.out "&nbsp;"
end sub
 
 
 '高亮显示搜索关键字
function searchH(body,key)
 searchh=replace(body,key,"<span style=""color:#ff00ff"">"&key&"</span>")  
end function 

'在增加表单中显示选择车间的表单
function formsscj()
   if session("level")=0 then 
	response.write"<select name='sscj' size='1'>"
    response.write"<option >请选择所属车间</option>"
    sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select>"  	 
   else 	 
     response.write"<input name='sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
     response.write"<input name='sscj' type='hidden' value="&session("level")&">"& vbCrLf
  end if 

end function


'百分之80表格显示
sub showpage_80(page,url,total,record,PgSz)
   response.write "<table width='80%' align='center'  border='0' cellpadding='0' cellspacing='0' class='border'><tr class='tdbg'><td><div align=""center"">"
   if page="" then page=1
   if page > 1 Then 
      response.write "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&PgSz="&PgSz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&PgSz="&PgSz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&PgSz="&PgSz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     response.write ""
   end if
   response.write"&nbsp;&nbsp;页次：<strong><font color=red>"&page&"</font>/"&total&"</strong>页&nbsp;&nbsp;"
  if Total =1 then 
    response.write"每页显示<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&PgSz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&PgSz='+this.value;"">条"
  else
   response.write"每页显示<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&PgSz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&PgSz='+this.value;"">条"
  end if 
   if Total =1 then 
    response.write"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&PgSz="&PgSz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    response.write"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&PgSz="&PgSz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 response.write"  <option value='"&page&"' selected >第"&page&"页</option>"
     else
    	 response.write"  <option value='"&ii&"'>第"&ii&"页</option>"
     end if 
   next 
   
   response.write" </select>&nbsp;&nbsp;共"&record&"条内容"
   response.write "</div></td></tr></table>"
end sub


'在增加表单中显示选择车间和班组的表单
function formsscjbz()
 dim rscj,sqlcj,rsbz,sqlbz
 if session("level")=0 then 
	'功能说明，先在levelname表中读取全部的levelclass=1的车间名称，然后根据车间ID在bzname表中读取对应的班组名称显示
	response.write"<select name='sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    response.write"<option  selected>选择所属车间</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	response.write"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    response.write"</select>"  	 & vbCrLf
    response.write "<select name='ssbz' size='1' >" & vbCrLf
    response.write "<option  selected>选择班组分类</option>" & vbCrLf
    response.write "</select></td></tr>  "  & vbCrLf
    response.write "<script><!--" & vbCrLf
    response.write "var groups=document.form1.sscj.options.length" & vbCrLf
    response.write "var group=new Array(groups)" & vbCrLf
    response.write "for (i=0; i<groups; i++)" & vbCrLf
    response.write "group[i]=new Array()" & vbCrLf
    response.write "group[0][0]=new Option(""选择班组分类"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0	
		sqlbz="SELECT * from bzname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   response.write "group["&rscj("levelid")&"][0]=new Option(""未添加班组"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'response.write"group["&rsbz("sscj")&"][0]=new Option(""车间"",""0"");" & vbCrLf
		   response.write"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("bzname")&""","""&rsbz("id")&""");" & vbCrLf
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
    response.write "var temp=document.form1.ssbz" & vbCrLf
    response.write "function redirect(x){" & vbCrLf
    response.write "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    response.write "temp.options[m]=null" & vbCrLf
    response.write "for (i=0;i<group[x].length;i++){" & vbCrLf
    response.write "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    response.write "}" & vbCrLf
    response.write "temp.options[0].selected=true" & vbCrLf
    response.write "}//--></script>" & vbCrLf
  else 	 
   response.write"<input name='sscj' type='text' value='"&sscjh(session("level"))&"'  disabled='disabled' >"& vbCrLf
   response.write"<input name='sscj' type='hidden' value="&session("level")&">"& vbCrLf
   sqlbz="SELECT * from bzname where sscj="&session("level")
   set rsbz=server.createobject("adodb.recordset")
   rsbz.open sqlbz,conn,1,1
   response.write"<select name='ssbz' size='1'>"
   
   if rsbz.eof and rsbz.bof then 
   	  response.write"<option value='0'>未添加班组</option>"
   else   
	  'response.write"<option value='0'>车间</option>"
      do while not rsbz.eof
	     response.write"<option value='"&rsbz("id")&"'>"&rsbz("bzname")&"</option>"
	  rsbz.movenext
      loop
	  end if 
	 response.Write"</select>" 
  rsbz.close
  set rsbz=nothing
 end if 
end function


'取当前网页URL
Function GetUrl() 
	'On Error Resume Next 
	Dim strtemp 
	If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
	 strtemp = "http://"
	Else 
	 strtemp = "https://"
	End If 
	strtemp = strtemp & Request.ServerVariables("SERVER_NAME") 
	If Request.ServerVariables("SERVER_PORT") <> 80 Then 
	 strtemp = strtemp & ":" & Request.ServerVariables("SERVER_PORT") 
	end if
	strtemp = strtemp & Request.ServerVariables("URL") 
	If Trim(Request.QueryString) <> "" Then 
	 strtemp = strtemp & "?" & Trim(Request.QueryString) 
	end if
	'判断URL中是否有分页函数，有则去掉
	if InStr(strtemp,"pgsz")<>0 then
		urllen=InStr(strtemp,"pgsz")
		strtemp=left(strtemp,urllen-2)
	end if  
	if InStr(strtemp,"pagelevelid")<>0 then
		urllen=InStr(strtemp,"pagelevelid")
		strtemp=left(strtemp,urllen-2)
	end if  
	'110725增加如果页面中有PAGE则删除掉
	if InStr(strtemp,"page")<>0 then
		urllen=InStr(strtemp,"page")
		strtemp=left(strtemp,urllen-2)
	end if  
	GetUrl = strtemp 
End Function


function message(text)

	dwt.out "<br/><br/><br/><div align='center'><DIV class='x-dlg x-dlg-closable x-dlg-draggable x-dlg-modal' style=' WIDTH: 263px; HEIGHT: 198px'>"
	dwt.out "  <DIV class='x-dlg-hd-left'>"
	dwt.out "    <DIV class='x-dlg-hd-right'>"
	dwt.out "      <DIV class='x-dlg-hd x-unselectable'>系统提示信息</DIV>"
	dwt.out "    </DIV>"
	dwt.out "  </DIV>"
	dwt.out "  <DIV class='x-dlg-dlg-body' style='WIDTH: 263px;'><div align=left>"
	dwt.out   text
	dwt.out "  </DIV></div>"
	dwt.out "</DIV></div>"
end function
'2008.10.16添加，在签名前加所属单位
public totalgr
Function usernameh2(userid,JJ,levelid)
    
	
	   dim sqlbz,rsbz
	if isnull(userid) then userid=0
	sqlbz="SELECT * from userid where levelid="&levelid&" and id="&userid
    set rsbz=server.createobject("adodb.recordset")
    rsbz.open sqlbz,conn,1,1
    if rsbz.eof then 
	
	else
	'do while not rsbz.eof
       	'dwt.out jj&"zzzzzzzzzzzz"&rsbz("groupid")&"<br>"
            totalgr=totalgr+1
		if jj=1 and rsbz("groupid")=10 then
			usernameh2=rsbz("username1")&"&nbsp;&nbsp;"
		end if
		if jj=2 and (rsbz("groupid")=10 or rsbz("groupid")=1 or rsbz("groupid")=4  or rsbz("groupid")=5  or rsbz("groupid")=6  or rsbz("groupid")=7  or rsbz("groupid")=8  or rsbz("groupid")=9  or rsbz("groupid")=24  or rsbz("groupid")=26) then
			usernameh2=rsbz("username1")&"&nbsp;&nbsp;"
		end if
		if jj=3 or isnull(jj) or jj=0 then   '111220修改添加 OR ISNULL 当VIEWDGROUP为空时也输出人名
			usernameh2=rsbz("username1")&"&nbsp;&nbsp;"
		end if
		
	'	rsbz.movenext
	'loop
	end if 
	rsbz.close
	set rsbz=nothing
end Function

Function usernameh3(j)
       dim sqlcj,rscj
	  sqlcj="SELECT * from levelname where levelid="&j
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	
	    usernameh3=replace(rscj("levelname"),"车间","")
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
end Function 
	'输出数据字典
	'selectname传递表单时用的名称
	'dicttitle数据字典中要调用的名称
	'onnumb当前选中的
	
	Public Function outdatadict(selectname,dicttitle,onnumb)
		outdatadict=""
		outdatadict="<select name='"&selectname&"' size='1'>"
		sqld="SELECT * from datadict where title='"&dicttitle&"'"
			set rsd=server.createobject("adodb.recordset")
			rsd.open sqld,connleft,1,1
			do while not rsd.eof
				outdatadict=outdatadict&"<option value='"&rsd("numb")&"'"
				if cint(onnumb)=rsd("numb") then outdatadict=outdatadict&" selected"
				outdatadict=outdatadict& ">"&rsd("info")&"</option>"& vbCrLf
				rsd.movenext
			loop
			rsd.close
			set rsd=nothing

		'outdatadict=outdatadict&"<option value='12'>12个月</option>"
		outdatadict=outdatadict&"</select>"
		'dwt.out outdatadict
	end function
Public Function outdatadict2(selectname,dicttitle,onnumb,jdzq)
		outdatadict2=""
		outdatadict2="<select name='"&selectname&"' size='1'>"
		sqld="SELECT * from datadict where title='"&dicttitle&"'"
			set rsd=server.createobject("adodb.recordset")
			rsd.open sqld,connleft,1,1
			do while not rsd.eof
	  
				outdatadict2=outdatadict2&"<option value='"&rsd("numb")&"'"
		if cint(jdzq)=rsd("numb") then outdatadict2=outdatadict2&" selected"
				outdatadict2=outdatadict2& ">"&rsd("info")&"</option>"& vbCrLf
				rsd.movenext				
			loop
			rsd.close
			set rsd=nothing

		outdatadict2=outdatadict2&"</select>"
	end function	
	Public Function dispalydatadict(dicttitle,onnumb)
		sqld="SELECT * from datadict where title='"&dicttitle&"' and numb="&cint(onnumb)
		set rsd=server.createobject("adodb.recordset")
		rsd.open sqld,connleft,1,1
		if rsd.eof and rsd.eof then 
			dwt.out "无"
		else
			dwt.out rsd("info")
		end if 	
	end function

%>