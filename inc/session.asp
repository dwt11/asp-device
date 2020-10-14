<%


'dim cUrl
'ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
'If ComeUrl = "" Then
'    Response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
   'Response.write "<br><br><br><p align=center><a href=/><font color='red'>请从首页登录</font></a></p>"
'Response.End
'end if



if session("UserName")="" then
response.write"<script>alert('您未登陆或页面停留时间过长，请重新登陆！');location='/'</script>"
response.end
end if
%>