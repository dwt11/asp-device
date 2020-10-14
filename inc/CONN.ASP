<%

   'on error resume next


'data.mdb 数据库包括技术档案、用户管理
db="/ybdata/data.mdb"
   
Set conn = Server.CreateObject("ADODB.Connection")
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
    conn.Open connstr


Sub CloseConn()
    On Error Resume Next
    If IsObject(conn) Then
        conn.Close
        Set conn = Nothing
    End If
    
End Sub
%>