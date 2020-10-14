<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<%
Dwt.out "<html>"& vbCrLf
Dwt.out "<head>"& vbCrLf
Dwt.out "<title>信息管理系统 >> 首页</title>"& vbCrLf
Dwt.out "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"& vbCrLf
Dwt.out "<link href='css/index.css' rel='stylesheet' type='text/css'> "& vbCrLf
Dwt.out "<script language=javascript>"& vbCrLf
Dwt.out "   <!--"& vbCrLf
Dwt.out "    function CheckForm() {"& vbCrLf
Dwt.out "      if(document.Login.UserName.value == '') {"& vbCrLf
Dwt.out "        alert('请输入用户名！');"& vbCrLf
Dwt.out "        document.Login.UserName.focus();"& vbCrLf
Dwt.out "        return false;"& vbCrLf
Dwt.out "      }"& vbCrLf
Dwt.out "      if(document.Login.password.value == '') {"& vbCrLf
Dwt.out "        alert('请输入密码！');"& vbCrLf
Dwt.out "        document.Login.password.focus();"& vbCrLf
Dwt.out "        return false;"& vbCrLf
Dwt.out "      }"& vbCrLf
Dwt.out "	  }"& vbCrLf
Dwt.out "  //-->"& vbCrLf
Dwt.out "</script>"& vbCrLf

Dwt.out "</head>"& vbCrLf
 
Dwt.out "<body leftmargin=0 topmargin=0>"& vbCrLf%>
<!--#include file="index_t.asp"-->

<%

DWT.OUT "<script type='text/javascript' src='images2006/qywh.js'></script> "&vbcrlf
Dwt.out "<table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>"& vbCrLf
Dwt.out "    <tr>"& vbCrLf
Dwt.out "     <td width=180 vAlign=top>"& vbCrLf
Dwt.out "      <!--用户登录代码开始-->"& vbCrLf
Dwt.out "        <table cellSpacing=0 cellPadding=0 width=""100%"" border=0>"& vbCrLf
Dwt.out "          <tr>"& vbCrLf
Dwt.out "            <td><IMG src=""/images2006/login_01.gif""></td>"& vbCrLf
Dwt.out "          </tr>"& vbCrLf
Dwt.out "          <tr>"& vbCrLf
Dwt.out "            <td vAlign=center align=middle background=/images2006/login_02.gif>"& vbCrLf
Dwt.out "<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0'>"& vbCrLf
Dwt.out "<form name='Login' action='login.asp' method='post' target='_parent'  onSubmit='return CheckForm();'>"& vbCrLf
Dwt.out "            <tr>"& vbCrLf
Dwt.out "                <td height='25' align='right'>用户名：</td><td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td>"& vbCrLf
Dwt.out "</tr>"& vbCrLf
Dwt.out "                <tr>"& vbCrLf
Dwt.out "                <td height='25' align='right'>密&nbsp;&nbsp;码：</td><td height='25'><input name='password' type='password' id='Password' size='10' maxlength='20'></td>"& vbCrLf
Dwt.out "                </tr>"& vbCrLf
Dwt.out "                <tr align='center'>"& vbCrLf
Dwt.out "                  <td height='37' colspan='2'>"& vbCrLf
Dwt.out "                         <input type='hidden' name='Action' value='Login'>"& vbCrLf
Dwt.out "		  <input type=""submit"" name=""Submit"" value=""登录"">"& vbCrLf
Dwt.out "&nbsp;&nbsp;<input name='Reset' type='reset' id='Reset' value=' 清除 '>"& vbCrLf
Dwt.out " </td>"& vbCrLf
Dwt.out "        </tr>"& vbCrLf
Dwt.out "            </form>"& vbCrLf
Dwt.out "		</table>"& vbCrLf
Dwt.out "        </td></tr><tr><td><IMG src=""/images2006/login_03.gif""></td></tr></table>"& vbCrLf
%>
<table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td class=left_title align=middle>通知公告</td>
          </tr>
          <tr>
            <td class=left_tdbg1 vAlign=top height=100>
			
<%i =0
sqlqxtb="SELECT top 6 * from xzgl_news where news_class=52 AND MONTH(NEWS_DATE)=MONTH(NOW()) AND YEAR(NEWS_DATE)=YEAR(NOW()) ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,conna,1,1
if rsqxtb.eof and rsqxtb.bof then 
'Dwt.out "<p align='center'>未添加新闻</p>" 
else
dwt.out "<marquee  scrollamount=2 height=100 onmouseover=stop()  onmouseout=start() direction='up'>"
do while not rsqxtb.eof
title=rsqxtb("news_title")
if len(title)>35 then
title=left(title,25)&"..."

%>
          
                <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target=_blank><%=title%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rsqxtb("news_date")%>]
                    <%else%>
                <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" target=_blank><%=rsqxtb("news_title")%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rsqxtb("news_date")%>]
                    <%end if
					i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
					dwt.out "</marquee>"
rsqxtb.close
set rsqxtb=nothing
if i>5 then Dwt.out "<Div align=right><a href=news_d.asp?NEWS_CLASS=52>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
             </td>
          </tr>
          <tr>
            <td class=left_tdbg2></td>
          </tr>
        </table>


<%



Dwt.out "</td>"& vbCrLf


Dwt.out "      <td width=5></td>"& vbCrLf
Dwt.out "      <td vAlign=top>"& vbCrLf

DWT.OUT "<DIV align='center'>真正的人才，不是能够评判是非、指出对错的人，因为几乎每一个人都能做到这一点<BR>真正的人才是能够让事情变得更好的人</DIV>"

Dwt.out "        <table cellSpacing=0 cellPadding=0 width=""100%"" border=0>"& vbCrLf
Dwt.out "          <tr>"& vbCrLf
Dwt.out "            <td class=main_title_1i>新 闻</td>"& vbCrLf
Dwt.out "          </tr>"& vbCrLf
Dwt.out "          <tr>"& vbCrLf
Dwt.out "            <td  vAlign=top class=main_tdbg_575>"& vbCrLf

'sqlnews="SELECT top 10 * from xzgl_news where xzgl_news_class.isindex=true ORDER BY news_date DESC"
sqlnews="SELECT top 11 xzgl_news.*,xzgl_news_class.isindex FROM xzgl_news INNER JOIN xzgl_news_class ON xzgl_news.news_class=xzgl_news_class.id WHERE (((xzgl_news_class.isindex)=True))  ORDER BY xzgl_news.news_date deSC"
set rsnews=server.createobject("adodb.recordset")
rsnews.open sqlnews,conna,1,1
if rsnews.eof and rsnews.bof then 
    Dwt.out "<p align='center'>未添加新闻</p>" 
else
    
iiii=0
	 '显示三日内党委工作动态
	 sqlbody1="SELECT top 2 * from dgtzl_body WHERE index=1 order by id desc"
	 set rsnews1=server.createobject("adodb.recordset")
	rsnews1.open sqlbody1,conndgt,1,1
	if rsnews1.eof and rsnews1.bof then 
	else
	  do while not rsnews1.eof
	  		iiii=iiii+1 
			   if DATEDIFF("d",rsnews1("news_date"),now()) <= 2 and iiii<3 then 
		   Dwt.out "<li><a href=/dw/view.asp?ID="&rsnews1("id")&" target=_blank>"&rsnews1("news_title")&"</a>&nbsp;&nbsp;[<a href='/dw/showlist.asp?classid=1'>党建动态</a>]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews1("news_date")&"]<br>"& vbCrLf
	  end if 
	  rsnews1.movenext
	  
	  loop
	end if 	 


do while not rsnews.eof
       dim i,addtime,nowtime
       title=rsnews("news_title")
	   addtime = rsnews("news_date") 
       nowtime=DATEDIFF("d",addtime,now())
       if len (title)>25 then
           title=left(title,25)&"..."
		   if DATEDIFF("d",addtime,now()) <= 3 then 
           Dwt.out "<li><span class=""classnew"">NEW: </span><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&title&"</a>&nbsp;&nbsp;[<a href='news_d.asp?NEWS_CLASS="&rsnews("news_class")&"'>"&newsclassh(rsnews("news_class"))&"</a>]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
		   else
		   Dwt.out "<li><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&title&"</a>&nbsp;&nbsp;[<a href='news_d.asp?NEWS_CLASS="&rsnews("news_class")&"'>"&newsclassh(rsnews("news_class"))&"]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"</a>]<br>"& vbCrLf
		   end if
       else
	       if DATEDIFF("d",addtime,now()) <= 3 then 
           Dwt.out "<li><span class=""classnew"">NEW: </span><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&rsnews("news_title")&"</a>&nbsp;&nbsp;[<a href='news_d.asp?NEWS_CLASS="&rsnews("news_class")&"'>"&newsclassh(rsnews("news_class"))&"</a>]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
		   else
		   Dwt.out "<li><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&rsnews("news_title")&"</a>&nbsp;&nbsp;[<a href='news_d.asp?NEWS_CLASS="&rsnews("news_class")&"'>"&newsclassh(rsnews("news_class"))&"</a>]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
		   end if
      end if
     i=i+1
	 
	 
	 
	 
'if i=8 then exit do
rsnews.movenext
loop
	
end if
rsnews.close
set rsnews=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_d.asp>更多新闻....</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Div>"
%>
            </td>
          </tr>
          <tr>
            <td class=main_shadow></td>
          </tr>
        </table>
      </td>
    </tr>
</table>







<table class=center_tdbgall cellSpacing=0 cellPadding=0 width=760 align=center border=0>
    <tr>
      <td class=left_tdbgall vAlign=top width=180>
<table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td class=left_title align=middle>每周缺陷检查</td>
          </tr>
          <tr>
            <td class=left_tdbg1 vAlign=top height=179>
			
<%i =0
sqlqxtb="SELECT distinct TOP 6 jcdate1,jcdate2 from mzqxdjzg where zgjg=false ORDER BY jcdate1 DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connb,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加</p>" 
else
do while not rsqxtb.eof
wzgnumb=connb.Execute("SELECT count(id) from mzqxdjzg where  zgjg=false and jcdate1=#"&rsqxtb("jcdate1")&"#")(0)
'dwt.out "33lk"&rsqxtb("jcdate1")&"k33"
%>
          
                <li><%dwt.out "<a href='mzqxdjzg_view.asp?jcdate1="&rsqxtb("jcdate1")&"&jcdate2="&rsqxtb("jcdate2")&"'  target=_blank>"&rsqxtb("jcdate1")&"到"&rsqxtb("jcdate2") &"未整改"&wzgnumb&"条</a>"%></li>
<%rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
%>
             </td>
          </tr>
          <tr>
            <td class=left_tdbg2></td>
          </tr>
        </table>	  
	  
	  
	  
	  
	  
	  
	  <table style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td class=left_title align=middle>缺陷整改通知</td>
          </tr>
          <tr>
            <td class=left_tdbg1 vAlign=top height=179>
			
<%i =0
sqlqxtb="SELECT top 6 * from scgl_qxtb ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connb,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加新闻</p>" 
else
do while not rsqxtb.eof

title=rsqxtb("qxtb_title")
if len (title)>35 then
title=left(title,25)&"..."

%>
          
                <li><a href="qxtb_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("qxtb_title")%>" target=_blank><%=title%></a>
                    <%else%>
                <li><a href="qxtb_view.asp?id=<%=rsqxtb("id")%>" target=_blank><%=rsqxtb("qxtb_title")%></a>
                    <%end if%>
                    <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
if i>5 then Dwt.out "<Div align=right><a href=qxtb_d.asp>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
             </td>
          </tr>
          <tr>
            <td class=left_tdbg2></td>
          </tr>
        </table>
  
  
  
  
  
  
  
  
  
  
  
  
        <table style="WORD-BREAK: break-all" cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">畅所欲言</td>
          </tr>
          <tr>
            <td class="left_tdbg1" valign="top" height="179">
<%i =0
sqlqxtb="SELECT top 8 * from csyy_body where news_class=1 ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,conncsyy,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rsqxtb.eof

	title=rsqxtb("news_title")
	if len (title)>25 then	title=left(title,25)&"..."
	%>
                  <li><a href="news_csyy_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target="_blank"><%=title%></a>
                  </li>                 
                      <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_csyy.asp?classid=1>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
</td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>

  
        <table style="WORD-BREAK: break-all" cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">黑板报</td>
          </tr>
          <tr>
            <td class="left_tdbg1" valign="top" height="179">
<%i =0
sqlqxtb="SELECT top 8 * from xzgl_news where news_class=41 ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connxzgl,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rsqxtb.eof

	title=rsqxtb("news_title")
	if len (title)>25 then	title=left(title,25)&"..."
	%>
                  <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target="_blank"><%=title%></a>
                  </li>                 
                      <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_d.asp?news_class=41>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
</td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        <table style="WORD-BREAK: break-all" cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">荣誉栏</td>
          </tr>
          <tr>
            <td class="left_tdbg1" valign="top" height="179"  id="container">
			<%
i =0
sqlqxtb="SELECT top 8 * from xzgl_news where news_class=42 ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connxzgl,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rsqxtb.eof

title=rsqxtb("news_title")
if len (title)>25 then
title=left(title,15)&"..."

%>

                  <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target="_blank"><%=title%></a>
                      <%else%>
                  </li>
                  <li><a href="news_view.asp?id=<%=rsqxtb("id")%>" target="_blank"><%=rsqxtb("news_title")%></a>
                      <%end if%>
                      <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_d.asp?news_class=42>更多内容....</a>&nbsp;&nbsp;</Div>"

			'下面内容为展示图片用
'			i =0
'DIM imgadd(10)
'sqlqxtb="SELECT top 5 * from xzgl_news where news_class=42 ORDER BY id DESC"
'set rsqxtb=server.createobject("adodb.recordset")
'rsqxtb.open sqlqxtb,connxzgl,1,1
'if rsqxtb.eof and rsqxtb.bof then 
''Dwt.out "<p align='center'>未添加内容</p>" 
'else
'do while not rsqxtb.eof
'	
'	leftnumb=instr(rsqxtb("news_body"),"Upload")-1
'  if leftnumb>0 then 	
'	strnumb=instr(rsqxtb("news_body"),".")-leftnumb
'	'response.write leftnumb&"ddddd"&strnumb
'	imgadd(i)=mid(rsqxtb("news_body"),leftnumb,strnumb+4)
'	
'	title=rsqxtb("news_title")
'	
'	
'	imgaddress=imgaddress&"{url:"""&imgadd(i)&""",link: ""news_view.asp?id="&rsqxtb("id")&"""},"
'	i=i+1
'  end if 	
''if i=8 then exit do
'rsqxtb.movenext
'loop
'end if
'rsqxtb.close
'set rsqxtb=nothing
'if i>7 then Dwt.out "<Div align=right><a href=qxtb_d.asp>更多内容....</a>&nbsp;&nbsp;</Div>"
'  IF LEFTNUMB>0 THEN imgaddress=mid(imgaddress,1,len(imgaddress)-1)
  	 
            ' response.write imgaddress&"]);"
	
%>

<script type="text/javascript">
<!--
/*    var o = new Rotate(165, 165, "container");
	o.addImg([<%=imgaddress%>]);
        o.show();
    function Rotate(width, height, container, timeout)
        {
                var isIE = navigator.appName == "Microsoft Internet Exploer";
                var isFF = /Firefox/i.test(navigator.userAgent);
                this.imgInfo = [];
                this.width = parseInt(width);
                this.height = parseInt(height);
                this.container = $(container);
                this.timeout = timeout || 3000;
                this.index = 0;
                this.oImg = null;
                this.timer = null;
                this.innerContainer = "asfman_" + uniqueID(6);
                this.order = "asfman_" + uniqueID(6);
                this.img = "asfman_" + uniqueID(6);
                this.template = "<div id='" + this.innerContainer + "'>\r\n" + "<div id='" + this.order + "'>\r\n{order}\r\n<div style='clear: both'></div>\r\n</div>";
                //add css 
                var styleCss = "#" + this.innerContainer + "{overflow: hidden; position: relative; width: " + this.width + "px; height: " + this.height + "px;}\r\n" +
                               "#" + this.order + "{position: absolute; right: 5px; bottom: 5px;}\r\n" +
                                           "#" + this.order + " a{width: 22px; line-height: 23px; height: 21px; font-size: 12px; text-align: center; margin: 3px 3px 0 3px; float: left;}\r\n" + "#" + this.order + " a:link, #" + this.order + " a:visited{background: transparent url(http://i3.sinaimg.cn/dy/deco/2007/1218/cc071218img/news_hdtj_ws_009.gif) no-repeat scroll 0pt 3px; color: #fff; text-decoration: none;}\r\n"+
                                           "#" + this.order + " a:hover{text-decoration:none;}\r\n" + 
                                           "#" + this.innerContainer + " img{border: 0; width: " + this.width + "px; height: " + this.height + "px;}\r\n";
                function uniqueID(n)
                {
                        var str="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
                        for(var ret = "", i = 0; i < n; i++)
                        {
                        ret += str.charAt(Math.floor(Math.random()*62));
                        }
                        return ret;
        };
                void function(cssText, doc)
                {
                        doc = doc || document;
                        var oStyle = doc.createElement("style");
                        oStyle.type = "text/css"; 
                        if(oStyle.styleSheet)
                        {
                                oStyle.styleSheet.cssText = cssText; 
                        }
                        else
                        {
                                oStyle.appendChild(doc.createTextNode(cssText));
                        } 
                        doc.getElementsByTagName("head")[0].appendChild(oStyle);
                }(styleCss);
        function $(str){return document.getElementById(str);};
                function addListener(o, type, fn)
                {
                        var func = function()
                        {
                          return function(){fn.call(o);}
                        }();
                        if(isIE)
                        {
                                o.attachEvent("on" + type, func);
                        }
                        else if(isFF)
                        {
                                o.addEventListener(type, func, false);
                        }else{
                                o["on" + type] = func;
                        }
                        return func;
                }
                if(Rotate.initialize == undefined)
                {
                   Rotate.prototype.addImg = function(obj)
                   {//{url: imgUrl, link: linkUrl, alt: alt, txt: txt}
                       if(obj)
                           {
                              if(obj.constructor == Array)
                                  {
                                     for(var i = 0, l =  obj.length; i < l; i++)
                                     {
                                         this.imgInfo.push(obj[i]);
                                     }
                                  }else
                                    this.imgInfo.push(obj);
                           }
                   }
                   Rotate.prototype.show = function()
                   {
                      var _this = this;
                          var order = "";
                          for(var i = 1, l =  this.imgInfo.length; i <= l; i++)
                      {
                          order += "<a href='javascript: void(0)'>" + i + "</a>";
                      }
              this.template = this.template.replace(/{order}/, order);
                          this.container.innerHTML = this.template;
                          for(var j = 0, len =  $(this.order).getElementsByTagName("a").length; j < len; j++)
                          {
                             addListener($(this.order).getElementsByTagName("a")[j], "click", clickFunc);
                                 addListener($(this.order).getElementsByTagName("a")[j], "focus", function(){this.blur();});
                          }
                          $(this.order).getElementsByTagName("a")[this.index].style.background = "transparent url(http://i3.sinaimg.cn/dy/deco/2007/1218/cc071218img/news_hdtj_ws_010.gif) no-repeat scroll 0pt 0px";
                          $(this.order).getElementsByTagName("a")[this.index].style.color = "#000";
                          function clickFunc()
                          {
                                  if(_this.timer)
                                         clearInterval(_this.timer);
                                  var n = (this.innerHTML.replace(/^\s+|\s+$/g,"")|0)-1;
                                  _this.change(n);
                                  var o = _this;
                                  o.index = n;
                                  o.timer = setInterval(function(){o.autoStart();}, o.timeout);
                          }
                          this.oA = document.createElement("a");
                          this.oA.href = "javascript: void(0)";
                          if(this.imgInfo[this.index].link)
                          {
                              this.oA.href = this.imgInfo[this.index].link
                                  this.oA.target = "_blank";
                          }
                          $(this.innerContainer).appendChild(this.oA);
                          this.oImg = new Image();
                          this.oImg.src = this.imgInfo[this.index].url;
                          this.oImg.style.filter = "revealTrans(duration=1,transition=23)";
                          this.oA.appendChild(this.oImg);
                          this.timer = setInterval(function(){_this.autoStart();}, this.timeout);

                   }
                   Rotate.prototype.autoStart = function()
                   {
                        this.index++;
                                var n = this.index >= this.imgInfo.length ? this.index = 0 : this.index;
                                this.change(n);
                   }
                   Rotate.prototype.change = function(index)
                   {
                          if(this.oImg.filters && this.oImg.filters.revealTrans)
                          {
                                 this.oImg.filters.revealTrans.Transition = Math.floor(Math.random()*23);
                                 this.oImg.filters.revealTrans.apply();
                                 this.oImg.src = this.imgInfo[index].url;
                                 this.oImg.filters.revealTrans.play();

                          }else
                            this.oImg.src = this.imgInfo[index].url;
                          this.oA.href = this.imgInfo[index].link;
                          for(var k = 0, length =  $(this.order).getElementsByTagName("a").length; k < length; k++)
                          {
                             $(this.order).getElementsByTagName("a")[k].style.background = "transparent url(http://i3.sinaimg.cn/dy/deco/2007/1218/cc071218img/news_hdtj_ws_009.gif) no-repeat scroll 0pt 3px;";
                                 $(this.order).getElementsByTagName("a")[k].style.color = "#fff";
                          }
                          $(this.order).getElementsByTagName("a")[index].style.background = "transparent url(http://i3.sinaimg.cn/dy/deco/2007/1218/cc071218img/news_hdtj_ws_010.gif) no-repeat scroll 0pt 0px";
                          $(this.order).getElementsByTagName("a")[index].style.color ="#000";
                   }
                   Rotate.initialize = true;
                }
    }
*/        
//-->
</script>

                
<%'IF LEFTNUMB>0 THEN 
'   DWT.OUT "<Div align=right><a href=news_d.asp?news_class=42>更多内容....</a>&nbsp;&nbsp;</div>"
'   ELSE
'     DWT.OUT "未添加内容"
'  end if 	 
   %>
              
                
                </td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>
        
                <table style="WORD-BREAK: break-all" cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="left_title" align="middle">生活小常识</td>
          </tr>
          <tr>
            <td class="left_tdbg1" valign="top" height="179"><%i =0
sqlqxtb="SELECT top 8 * from xzgl_news where news_class=43 ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connxzgl,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rsqxtb.eof

title=rsqxtb("news_title")
if len (title)>25 then
title=left(title,25)&"..."

%>
                  <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target="_blank"><%=title%></a>
                      <%else%>
                  </li>
                  <li><a href="news_view.asp?id=<%=rsqxtb("id")%>" target="_blank"><%=rsqxtb("news_title")%></a>
                      <%end if%>
                      <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
rsqxtb.close
set rsqxtb=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_d.asp?news_class=43>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
                  </li>
             </td>
          </tr>
          <tr>
            <td class="left_tdbg2"></td>
          </tr>
        </table>

        
        
        
        
        
        </td>
      <td width=5></td>
      <td vAlign=top>
        
        
        
        
        
        
        <table cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tr>
            <td>
              
              
              
              
              
              
              
              
              
              
              <table cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=main_title_1i>车间未整改缺陷统计</td>
                </tr>
                <tr>
                  <td class=main_tdbg_282i vAlign=top height=136>
                    <%i =0
sqlqxdj="SELECT top 15 * from qxdjzg where zgjg=false ORDER BY id DESC"
set rsqxdj=server.createobject("adodb.recordset")
rsqxdj.open sqlqxdj,connb,1,1
if rsqxdj.eof and rsqxdj.bof then 
    Dwt.out "<p align='center'>未添加内容</p>" 
else
   do while not rsqxdj.eof

title=rsqxdj("body")
'if len (title)>13 then
'  title=left(title,10)&"..."
'end if 
  Dwt.out "<li>"&sscjh_d(rsqxdj("sscj"))&"&nbsp;"&rsqxdj("wh")&"&nbsp;&nbsp;"&title&"&nbsp;&nbsp;督办人："&rsqxdj("dbname")&"</li>"
  rsqxdj.movenext
loop
end if
rsqxdj.close
set rsqxdj=nothing
'Dwt.out "<br>请登陆后查看详细内容"
%>
                 </td>
                </tr>
              </table>
              
              
              
              
              
              
              <table cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=main_title_1i>工作完成情况</td>
                  <td class=main_title_1i width="20%"><a href=diaoduhui_d.asp?wangong=2><SPAN>更多...</SPAN></td>
                </tr>
                <tr>
                  <td height=136 colspan="2" vAlign=top class=main_tdbg_282i>
                   <%dwt.out "<marquee  scrollamount=2 height=200 onmouseover=stop()  onmouseout=start() direction='up'>"

			
i =0
sqlqxtb="SELECT  * from huiyiluoshi where isno=false ORDER BY id DESC"
'sqlqxtb="SELECT top 7 * from huiyiluoshi where pxst_class=1 and isno=false ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,connpxjhzj,1,1
if rsqxtb.eof and rsqxtb.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rsqxtb.eof

	title=rsqxtb("pxst_title")
	if len (title)>25 then	title=left(title,25)&"..."
	%>
<li><a href="diaoduhui_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("pxst_title")%>" target="_blank" style="color:#FF0000"><%=title%>&nbsp;&nbsp;&nbsp;&nbsp;责任单位及责任人：<%=rsqxtb("zr_danwei")%>&nbsp;&nbsp;<%=rsqxtb("zr_ren")%>&nbsp;&nbsp;&nbsp;&nbsp;未完成</a></li>                 
                      <%i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
dwt.out "</marquee>"

rsqxtb.close
set rsqxtb=nothing


%>
                 </td>
                </tr>
              </table>
              
              
              
              
              
              
              
              
              
              
              
              
              
              
              
			  
			  
			  <table cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=main_title_1i>科技信息</td>
                </tr>
                <tr>
                  <td class=main_tdbg_282i vAlign=top height=136>
                    <%i =0
sqlnews="SELECT top 10 * from xzgl_news where news_class=11 ORDER BY news_date desc,id DESC"
set rsnews=server.createobject("adodb.recordset")
rsnews.open sqlnews,conna,1,1
if rsnews.eof and rsnews.bof then 
    Dwt.out "<p align='center'>未添加新闻</p>" 
else
    do while not rsnews.eof
       
       title=rsnews("news_title")
       if len (title)>25 then
           title=left(title,25)&"..."
           Dwt.out "<li><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&title&"</a>&nbsp;&nbsp;["&newsclassh(rsnews("news_class"))&"]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
       else
           Dwt.out "<li><a href=news_view.asp?ID="&rsnews("id")&"&title="&rsnews("news_title")&" target=_blank>"&rsnews("news_title")&"</a>&nbsp;&nbsp;["&newsclassh(rsnews("news_class"))&"]&nbsp;&nbsp;&nbsp;&nbsp;["&rsnews("news_date")&"]<br>"& vbCrLf
      end if
     i=i+1
'if i=8 then exit do
rsnews.movenext
loop
end if
rsnews.close
set rsnews=nothing
if i>7 then Dwt.out "<Div align=right><a href=news_d.asp?NEWS_CLASS=11>更多科技信息....</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Div>"
%>
                    </ul></td>
                </tr>
              </table>
              <table cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=main_title_1i>培训试题</td>
                </tr>
                <tr>
                  <td class=main_tdbg_282i vAlign=top height=136>
                    
                     <%i =0
sqlpxst="SELECT top 8 * from pxst ORDER BY id DESC"
set rspxst=server.createobject("adodb.recordset")
rspxst.open sqlpxst,conne,1,1
if rspxst.eof and rspxst.bof then 
Dwt.out "<p align='center'>未添加内容</p>" 
else
do while not rspxst.eof

title=rspxst("pxst_title")
if len (title)>25 then
title=left(title,25)&"..."

%>
            
                <li><a href="pxst_view.asp?ID=<%=rspxst("id")%>" title="<%=rspxst("pxst_title")%>" target=_blank><%=title%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rspxst("pxst_date")%>]<br>
                    <%else%>
                <li><a href="pxst_view.asp?id=<%=rspxst("id")%>" target=_blank><%=rspxst("pxst_title")%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rspxst("pxst_date")%>]<br>
                    <%end if%>
                    <%i=i+1
'if i=8 then exit do
rspxst.movenext
loop
end if
rspxst.close
set rspxst=nothing
if i>7 then Dwt.out "<Div align=right><a href=pxst_d.asp>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
                    </ul></td>
                </tr>
              </table>
            <!--频道一最新文章代码结束--></td>
            <td vAlign=top width=4></td>
          </tr>
          <tr>
            <td vAlign=top>
            <!--专题一最新文章代码开始-->
              <table cellSpacing=0 cellPadding=0 width="100%" border=0>
                <tr>
                  <td class=main_title_1i>值班表</td>
                </tr>
                <tr>
                  <td class=main_tdbg_282i vAlign=top height=136>
		<ul>		              
                     <%i =0
sqlzbb="SELECT top 8 * from zbb ORDER BY id DESC"
set rszbb=server.createobject("adodb.recordset")
rszbb.open sqlzbb,conna,1,1
if rszbb.eof and rszbb.bof then 
Dwt.out "<p align='center'>未添加新闻</p>" 
else
do while not rszbb.eof

%>
              
               
                <li><a href="zbb_view.asp?id=<%=rszbb("id")%>" target=_blank><%=rszbb("title")%></a></LI>
                    
                    <%i=i+1
'if i=8 then exit do
rszbb.movenext
loop
end if
rszbb.close
set rszbb=nothing
if i>7 then Dwt.out "<Div align=right><a href=zbb_d.asp>更多内容....</a>&nbsp;&nbsp;</Div>"
%>
                            </ul></td>
                </tr>
              </table>
            <!--专题一最新文章代码结束--></td>
            <td width=4></td>
          </tr>
          <tr>
            <td class=main_shadow colSpan=2></td>
          </tr>
          
          <tr>
            <td class=main_shadow colSpan=2></td>
          </tr>
        </table>
      </td>
    </tr>
</table>
  <TABLE width=760 border=0 align="center" cellPadding=0 cellSpacing=0 
background=images2006/bottom_back.gif>
    <TBODY>
      <TR>
        <TD class=sxpta-font2 align=middle height=24>设备管理系统</TD>
        <TD width=140 height=54 rowSpan=2><IMG height=54 
      src="images2006/bottom_r.gif" width=140 useMap=#Map 
  border=0></TD>
      </TR>
      <TR>
        <TD class=sxpta-font2 align=middle height=30>
          <TABLE class=black cellSpacing=0 cellPadding=0 width=610 align=center 
      border=0>
            <TBODY>
              <TR>
                <TD width=170> </TD>
                <TD vAlign=bottom width=394 height=28> </TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      </TR>
    </TBODY>
  </TABLE>
</body>
</html>