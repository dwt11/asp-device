<!--#include file="conn.asp"-->
 <link rel="stylesheet" href="style.css" type="text/css">




 <style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
 </style>
<%if request("action")="save"  then 
 set rs4=Server.createObject("adodb.recordset")
         strsql="SELECT  * from moban where id="&request("id")
         rs4.open strsql,conn,1,3
		 rs4("moban")=request.form(request("id"))
				 rs4.update
		 rs4.close
		 response.write"<Script Language=Javascript>window.alert('����ģ��ɹ�');location.href('moban.asp?id="&request("id")&"')</Script>"
else  %>
  	  
	  
	  
	   <form name="form1" method="post" action="bugpost.asp?action=save">

   <div align="center">
     
	
	
  
	  <textarea name="body" cols="40" rows="20"></textarea>
	  
	   
  </div>


   
   
   

     <div align="center">
       <input type="submit" name="Submit" value="�ύ">
     </div>
	   </form>

 
<%end if %>
	   
