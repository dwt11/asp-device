<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Server.ScriptTimeOut=5000
%>
<!--#include file="UpLoadClass.asp"-->
<%
if Request.QueryString("action")="upload" then
dim myrequest,lngUpSize
Set myrequest=new UpLoadClass
lngUpSize = myrequest.Open()
  select case myrequest.error
         case 0
		 response.Write("<script>window.parent.LoadIMG('"&myrequest.savepath&myrequest.form("file1")&"');</script>")
         case 1
		 response.Write("<script>alert('�ļ�����');window.parent.$('divProcessing').style.display='none';history.back();</script>")
		 case 2
		 response.Write("<script>alert('�������ϴ������͵��ļ���');window.parent.$('divProcessing').style.display='none';history.back();</script>")
                 case 3
		 response.Write("<script>alert('�������ϴ������͵��ļ���');window.parent.$('divProcessing').style.display='none';history.back();</script>")
		 case else
		 response.Write("<script>alert('�ļ��ϴ�����');window.parent.$('divProcessing').style.display='none';history.back();</script>")
  end select
end if
%>