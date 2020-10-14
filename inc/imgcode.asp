<%
function imgCode(strContent)
dim re
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
		
re.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
strContent=re.replace(strContent,"<div align=center><img SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>550)this.width=333""></div>")


set re=Nothing
imgCode=strContent
end function

%>
