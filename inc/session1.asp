<%


dim cUrl
ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
 

if session("UserName")="" then
response.write"<script>alert('��δ��½���½����ҳ�������');location='/'</script>"
response.end
end if
%>