<%







Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName,ConnStr,Conn 
SqlDatabaseName = "qchgl" 'SQL数据库名 
SqlUsername = "qch" 'SQL数据库用户名 
SqlPassword = "123456" 'SQL数据库用户密码 
SqlLocalName = "172.16.10.129" 'SQL主机IP地址（服务器名） 
'SqlUsername = "sa" 'SQL数据库用户名 
'SqlPassword = "admin" 'SQL数据库用户密码 
'SqlLocalName = "(Local)\GSQL" 'SQL主机IP地址（服务器名） 
ConnStr = " Password = "&SqlPassword&"; user id =" & SqlUsername&"; Initial Catalog ="&SqlDatabaseName&"; data source ="& SqlLocalName &";Provider = sqloledb;" 
Set connjlhs = Server.CreateObject("ADODB.Connection") 
connjlhs.open ConnStr 

If connjlhs.state=2 Then 
'esponse.Write("正在连接") 
ElseIf connjlhs.state=1 Then 
'Response.Write("连接已经打开") 
Else 
'Response.Write("对不起，无法打开连接") 
End If 


%>