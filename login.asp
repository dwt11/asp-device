<%@language=vbscript codepage=936%>
<%
Option Explicit%>
<!--#include file="conn.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
Action = Trim(Request("Action"))
If Action = "Login" Then
    Call ChkLogin
ElseIf Action = "Logout" Then
    Call Logout
	else
	   response.write"<Script Language=Javascript>window.alert('����ֱ�Ӵ򿪴�ҳ��');location.href='index.asp';</Script>"
End If
Call CloseConn


Sub ChkLogin()
    Dim sql, rs,level
    Dim UserName,UserName1, Password, CheckCode, RndPassword, logincode,trueip
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    Password = ReplaceBadChar(Trim(Request("Password")))
    Password=md5(Password,16)
	Set rs = Server.CreateObject("adodb.recordset")
    sql = "select * from userid where Password='" & Password & "' and userName='" & UserName & "'"
    rs.Open sql, Conn, 1, 3
    If rs.bof And rs.EOF Then
          response.write"<Script Language=Javascript>window.alert('�û������������!');location.href='index.asp';</Script>"
         Exit Sub
      End If
    if rs("levelzclass")=0 then 
		level=conn.Execute("SELECT levelclass FROM levelname WHERE levelid="&rs("levelid"))(0)
		session("level")=level
    else
	    session("level")=3
	end if 	
	
	session("UserName")=rs("UserName")  '�û���¼��
    session("UserName1")=rs("UserName1")  '�û���ʵ����
    'session("pageleveltext")=rs("pagelevel")   '�û���ҳ��Ȩ��
    'session("groupleveltext")=rs("grouplevel")   '�û�����Ȩ��
    session("groupid")=rs("groupid")
	session.Timeout=1000
	session("levelclass")=rs("levelid")       '�û�������(����)
	session("levelzclass")=rs("levelzclass")  '�û���������¼�(����)
    session("userid")=rs("id")        '�û���id
    TrueIP = Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
    If TrueIP = "" Then TrueIP = Request.ServerVariables("REMOTE_ADDR")
	rs("dldate") = Now()      '��¼ʱ��
    rs("dlcs") = rs("dlcs") + 1  '��¼����
    rs("dlip")=TrueIP    '��¼IP
	rs.Update
	rs.Close
	dwt.savesl "","","��¼�ɹ�"
    Call CloseConn
    Response.Redirect "main.asp"
End Sub

Sub Logout()
	dwt.savesl "","","�˳���¼"
    session("Level")=""
    session("UserName")=""
	session("userid")=""
	session("pagelevelid")=""    '��ǰ�򿪵�ҳ��ID����IDȡ��left.mdb����LEFT.ASP������
	session("pageleveltext")=""
	session("groupleveltext")=""
	session("jcdate1")=""    '���ÿ��ȱ�ݼ�¼ʱ�� MZQXDJZG.ASP
	session("jcdate2")=""
    Call CloseConn
    Response.Redirect "/"
End Sub

%>