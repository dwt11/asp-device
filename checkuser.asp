<!--#include file="conn.asp"-->

<%
	username = unescape(Request("username"))
	sql = "select top 1 * from userid where username='"&username&"'"
	rs = Conn.Execute(sql)
	sResult=rs("username")
    Response.Write(escape(sResult))


Call CloseConn
%>