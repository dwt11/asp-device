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
	   response.write"<Script Language=Javascript>window.alert('不能直接打开此页面');location.href='index.asp';</Script>"
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
          response.write"<Script Language=Javascript>window.alert('用户名或密码错误!');location.href='index.asp';</Script>"
         Exit Sub
      End If
    if rs("levelzclass")=0 then 
		level=conn.Execute("SELECT levelclass FROM levelname WHERE levelid="&rs("levelid"))(0)
		session("level")=level
    else
	    session("level")=3
	end if 	
	
	session("UserName")=rs("UserName")  '用户登录名
    session("UserName1")=rs("UserName1")  '用户真实姓名
    'session("pageleveltext")=rs("pagelevel")   '用户的页面权限
    'session("groupleveltext")=rs("grouplevel")   '用户的组权限
    session("groupid")=rs("groupid")
	session.Timeout=1000
	session("levelclass")=rs("levelid")       '用户所属组(车间)
	session("levelzclass")=rs("levelzclass")  '用户所属组的下级(班组)
    session("userid")=rs("id")        '用户的id
    TrueIP = Trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
    If TrueIP = "" Then TrueIP = Request.ServerVariables("REMOTE_ADDR")
	rs("dldate") = Now()      '登录时间
    rs("dlcs") = rs("dlcs") + 1  '登录次数
    rs("dlip")=TrueIP    '登录IP
	rs.Update
	rs.Close
	dwt.savesl "","","登录成功"
    Call CloseConn
    Response.Redirect "main.asp"
End Sub

Sub Logout()
	dwt.savesl "","","退出登录"
    session("Level")=""
    session("UserName")=""
	session("userid")=""
	session("pagelevelid")=""    '当前打开的页面ID，此ID取自left.mdb，在LEFT.ASP中设置
	session("pageleveltext")=""
	session("groupleveltext")=""
	session("jcdate1")=""    '添加每周缺陷记录时用 MZQXDJZG.ASP
	session("jcdate2")=""
    Call CloseConn
    Response.Redirect "/"
End Sub

%>