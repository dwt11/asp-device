<%




   Set fs   =   server.CreateObject("scripting.filesystemobject")   
   set myfile   =   fs.CreateTextFile("11.xls",true)   
   Set rs   =   Server.CreateObject("ADODB.Recordset")   
   myfile.writeline strLineTM   
   rs.Close   
   set rs=nothing   



%>