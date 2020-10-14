<!--#include file="conn.asp"-->

<%
	qptjtz_bh =unescape(request("qptjtz_bh"))
	sql = "select * from qptjtz where bh='"&qptjtz_bh&"'"
	rs = Connjg.Execute(sql)
	sResult=rs("qptzid")
    Response.Write(escape(sResult))


Call CloseConn
%>