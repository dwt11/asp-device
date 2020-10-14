<%@ CODEPAGE=65001 %>  
<% Response.CodePage=65001%>  
<% Response.Charset="UTF-8" %>
<!--#include file="../conn.asp"-->
<!-- #INCLUDE FILE="Includes/FusionCharts.asp" -->
<!-- #INCLUDE FILE="Includes/Functions.asp" -->
<!-- #INCLUDE FILE="DataGen.asp" -->
<%

'分类统计点击后,返回分类的中文名称到JS
if request("action")="getclassname" then strResult=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&request("dclass"))(0)

'车间统计点击后,返回车间的中文名称到JS
if request("action")="getsscjname" then strResult=conn.Execute("SELECT levelname FROM levelname WHERE levelid="&request("sscj"))(0)

Response.Write(strResult)
%>






