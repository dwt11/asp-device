<%


dim cUrl
ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
 

if session("UserName")="" then
response.write"<script>alert('您未登陆请登陆后在页面输出！');location='/'</script>"
response.end
end if
%>