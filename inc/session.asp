<%


'dim cUrl
'ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
'If ComeUrl = "" Then
'    Response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
   'Response.write "<br><br><br><p align=center><a href=/><font color='red'>�����ҳ��¼</font></a></p>"
'Response.End
'end if



if session("UserName")="" then
response.write"<script>alert('��δ��½��ҳ��ͣ��ʱ������������µ�½��');location='/'</script>"
response.end
end if
%>