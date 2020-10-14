<%@ CODEPAGE=65001 %>  
<% Response.CodePage=65001%>  
<% Response.Charset="UTF-8" %>
<!--#include file="../conn.asp"-->
<!-- #INCLUDE FILE="Includes/FusionCharts.asp" -->
<!-- #INCLUDE FILE="Includes/Functions.asp" -->
<!-- #INCLUDE FILE="DataGen.asp" -->
<%











'以下从JS中过来的连接接收数据,点击饼图后显示
	Dim intYear
	intYear = Request.QueryString("Year")
	if intYear="" then
		intYear = year(date())
	end if
'This method writes the employee yearly sales data as XML.
'To this page, we're provided employeed Id.
Dim eId
eId = Request.QueryString("id")
'XML Data container
Dim strXML
		  '显示月检修量
		 '中文乱码 strclassmonthXML  = "<chart caption='"&intYear&"年 "&sbclass_name&" 每月检修量'  palette='" & getPalette() & "' animation='" & getAnimationState()& "' subcaption='' formatNumberScale='0' numberPrefix='' showValues='0' seriesNameInToolTip='0'>"
		
		'引处无法传递中文名称,所以标题,用GETNAME.ASP来传递
		 strclassmonthXML  = "<chart caption=''  palette='" & getPalette() & "' animation='" & getAnimationState()& "' subcaption='' formatNumberScale='0' numberPrefix='' showValues='0' seriesNameInToolTip='0'>"
		  strclassmonthXML= strclassmonthXML& getsscjBymonth(intYear,eId)
		  strclassmonthXML =strclassmonthXML& "<styles><definition><style type='font' color='" & getCaptionFontColor() & "' name='CaptionFont' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
		  strclassmonthXML =strclassmonthXML &"</chart>"

'Output it
Response.ContentType = "text/xml"
Response.Write(strclassmonthXML)
%>






