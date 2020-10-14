<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/session.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/sb_function.asp"-->


<!-- #INCLUDE FILE="Includes/FusionCharts.asp" -->
<!-- #INCLUDE FILE="Includes/Functions.asp" -->
<!-- #INCLUDE FILE="Includes/PageLayout.asp" -->
<!-- #INCLUDE FILE="DataGen.asp" -->





<html>
<head>
<title>统计汇总</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='../css/ext-all.css' rel='stylesheet' type='text/css'>
<link href='../css/body.css' rel='stylesheet' type='text/css'>


	<script LANGUAGE="Javascript" SRC="FusionCharts/FusionCharts.js"></script>		
	<script LANGUAGE="JavaScript">				
	var dclassChartLoaded=false;
	
	
	function FC_Rendered(DOMId){
		switch(DOMId){
			
			//分类的柱图在页面中加载好了
			case "dclassDetails":				
				dclassChartLoaded = true;
				break;
			//车间的柱图在页面中加载好了
			case "sscjDetails":				
				sscjChartLoaded = true;
				break;
//			case "SalesByProd":				
//				prodChartLoaded = true;
//				break;
		}
		return;
	}


    //点击分类饼图画,更新分类的详细月信息
	function updateChartClass(dclass,intyear){			
		//Update the chart only if has loaded
		if (dclassChartLoaded){
			  //引处无法传递中文名称,所以标题,用GETNAME.ASP来传递
			  var oBao = false;
			  if (!oBao && typeof XMLHttpRequest != 'undefined') {
					oBao = new XMLHttpRequest();
				  }
				  //特殊字符：+,%,&,=,?等的传输解决办法.字符串先用escape编码的.
				  //Update:2004-6-1 12:22
					  var userInfoo = "dclass="+dclass;
					  oBao.open("POST","getname.asp?action=getclassname",false);
					  oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
					  oBao.send(userInfoo);
				  //服务器端处理返回的是经过escape编码的字符串.
				  var strResult = unescape(oBao.responseText);			
			
				  jxnrgh.innerHTML="<b>"+strResult+intyear+"年每月检修量</b>";
			//alert(unescape(name1));
			//DataURL for the chart
			//var strURL = "Data_classByCategory.asp?id=" + dclass + "&name=" +name+ "&intyear=" +intyear;
			var strURL = "Data_classByCategory.asp?id=" + dclass + "&year=" +intyear;
			//var strURL = "Data_classByCategory.asp?id=" + dclass + "&intyear=" +intyear;		
			//Sometimes, the above URL and XML data gets cached by the browser.
			//If you want your charts to get new XML data on each request,
			//you can add the following line:
			strURL = strURL + "&currTime=" + getTimeForURL();
			//getTimeForURL method is defined below and needs to be included
			//This basically adds a ever-changing parameter which bluffs
			//the browser and forces it to re-load the XML data every time.
								
			//URLEncode it - NECESSARY.
			//strURL = escape(strURL);
		
			//Get reference to chart object using Dom ID "EmployeeDetails"
			var chartObj = getChartFromId("dclassDetails");			
			//Send request for XML
			chartObj.setDataURL(strURL);
		} else {
			//Show error
			alert("正在加载请等待.");
			return;
		}
		
		
	}




    //点击车间饼图后,更新车间的详细月信息
	function updateChartSscj(sscj,intyear){			
		//Update the chart only if has loaded
		if (sscjChartLoaded){
			  //引处无法传递中文名称,所以标题,用GETNAME.ASP来传递
			  var oBao = false;
			  if (!oBao && typeof XMLHttpRequest != 'undefined') {
					oBao = new XMLHttpRequest();
				  }
				  //特殊字符：+,%,&,=,?等的传输解决办法.字符串先用escape编码的.
				  //Update:2004-6-1 12:22
					  var userInfoo = "sscj="+sscj;
					  oBao.open("POST","getname.asp?action=getsscjname",false);
					  oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
					  oBao.send(userInfoo);
				  //服务器端处理返回的是经过escape编码的字符串.
				  var strResult = unescape(oBao.responseText);			
			
				  jxnrgh.innerHTML="<b>"+strResult+intyear+"年每月检修量</b>";
			//alert(unescape(name1));
			//DataURL for the chart
			//var strURL = "Data_classByCategory.asp?id=" + dclass + "&name=" +name+ "&intyear=" +intyear;
			var strURL = "Data_sscjByCategory.asp?id=" + sscj + "&year=" +intyear;
			//var strURL = "Data_classByCategory.asp?id=" + dclass + "&intyear=" +intyear;		
			//Sometimes, the above URL and XML data gets cached by the browser.
			//If you want your charts to get new XML data on each request,
			//you can add the following line:
			strURL = strURL + "&currTime=" + getTimeForURL();
			//getTimeForURL method is defined below and needs to be included
			//This basically adds a ever-changing parameter which bluffs
			//the browser and forces it to re-load the XML data every time.
								
			//URLEncode it - NECESSARY.
			//strURL = escape(strURL);
		
			//Get reference to chart object using Dom ID "EmployeeDetails"
			var chartObj = getChartFromId("sscjDetails");			
			//Send request for XML
			chartObj.setDataURL(strURL);
		} else {
			//Show error
			alert("正在加载请等待.");
			return;
		}
		
		
	}



	function getTimeForURL(){
		var dt = new Date();
		var strOutput = "";
		strOutput = dt.getHours() + "_" + dt.getMinutes() + "_" + dt.getSeconds() + "_" + dt.getMilliseconds();
		return strOutput;
	}
	function openNewWindow(theURL,winName,features) {
		 window.open(theURL + "?year=" + getSelectedYear(),winName,features);
	}
	function getSelectedYear(){
		var selYear;
		for (i=0; i<document.frmYr.Year.length; i++){			
			if(document.frmYr.Year[i].checked){				 
				selYear = document.frmYr.Year[i].value;
			}
		}
		return selYear;
	}
	</script>


</head>
<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>


	<div style='left:6px;'>
	     <DIV class='x-layout-panel-hd'>
	        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'><a href='sb_fx.asp?Year=<%=Request.Form("Year")%>'><b>总量统计</b></a>&nbsp;&nbsp;&nbsp;<a href='?action=totalclass&Year=<%=Request.Form("Year")%>'>按分类统计</a>&nbsp;&nbsp;&nbsp;<a href='?action=totalsscj&Year=<%=Request.Form("Year")%>'>按车间统计</a>
            
            
            
            &nbsp;&nbsp;&nbsp;<a href='?action=totalssbz&Year=<%=Request.Form("Year")%>'>按班组统计</a>            
			
			
            </span>
	     </div>

	<div class='x-toolbar'><div align=left>

    
<%
	Dim oRs, strSQL
	
	Dim intYear
	intYear = Request("Year")
	
	if intYear="" then
		intYear = year(date())
	end if	
	
	Dim animateCharts
	animateCharts = Request.Form("animate")
	if animateCharts="" then
		animateCharts = "1"
	end if
	Session("animation") = animateCharts
'显示年选择框
	Call render_yearSelectionFrm(request("action"))
	
%>
	
	
	</div></div>
<%
if request("action")="" then call main
if request("action")="totalclass" then call totalclass
if request("action")="totalsscj" then call totalsscj
if request("action")="totalssbz" then call totalssbz



sub main()
  
  
  
  %>
  
  
		<Div class='x-layout-panel' style='WIDTH: 100%;'>
		<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
		<tr class="x-grid-header">
		     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>
		    </tr>
			  <tr class='x-grid-row x-grid-row-alt' >
    <td align="center"  bgcolor="#EDF9D5" >
	<%	
	
	'显示年检修量
	strYearXML  = "<chart caption='所有年检修量' palette='" & getPalette() & "' animation='" & getAnimationState()& "' subcaption='' formatNumberScale='0' numberPrefix='' showValues='0' seriesNameInToolTip='0'>"
	strYearXML= strYearXML& getSalesByYear()
	strYearXML =strYearXML& "<styles><definition><style type='font' color='" & getCaptionFontColor() & "' name='CaptionFont' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
	strYearXML =strYearXML &"</chart>"
	call  renderChart("FusionCharts/MSColumn3DLineDY.swf", "",strYearXML,"SalesByYear", 500, 325, false, true)
	%>
      
      </td></tr>
			<tr class='x-grid-row x-grid-row-alt' >
    <td align="center"  bgcolor="#EDF9D5" >
	<%	
	
	'显示月检修量
	strmonthXML  = "<chart caption='"&intYear&"年每月检修量'  palette='" & getPalette() & "' animation='" & getAnimationState()& "' subcaption='' formatNumberScale='0' numberPrefix='' showValues='0' seriesNameInToolTip='0'>"
	strmonthXML= strmonthXML& getSalesBymonth(intYear)
	strmonthXML =strmonthXML& "<styles><definition><style type='font' color='" & getCaptionFontColor() & "' name='CaptionFont' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
	strmonthXML =strmonthXML &"</chart>"
	call  renderChart("FusionCharts/MSColumn3DLineDY.swf", "",strmonthXML,"SalesBymonth", 500, 325, false, true)
	%>
      </td></tr>
		</table> <%
end sub	





sub totalclass()
'按分类统计  %>
  
  
		<Div class='x-layout-panel' style='WIDTH: 100%;'>
		<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
		<tr class="x-grid-header">
		     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>
		    </tr>



</tr>
  <tr class='x-grid-row x-grid-row-alt' >  <td  bgcolor="#EDF9D5"  align="center"  >
    
    <%
	'显示分类检修量年比例 饼图
				strEmployeeXML  = "<chart caption='" &  intYear &  "年分类检修比例' subcaption='点击饼图中的分类,显示每月分类工作量' palette='" & getPalette()&  "' animation='" &  getAnimationState()&  "'  showValues='0' numberPrefix='' formatNumberScale='0' showPercentInToolTip='0'>"
				strEmployeeXML =  strEmployeeXML&getSalePerEmpXML(intYear,-1,false,true,false)
				strEmployeeXML =  strEmployeeXML&"<styles><definition><style type='font' name='CaptionFont' color='"&  getCaptionFontColor()&  "' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
				strEmployeeXML = strEmployeeXML& "</chart>"
				call renderChart("FusionCharts/Pie3D.swf","",strEmployeeXML, "TopEmployees", 600, 425, false, false)
				%>
    </td>
  </tr>
 
 
	
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
      <tr class='x-grid-row x-grid-row-alt' >  <td  bgcolor="#EDF9D5"  align="center"  >
      
 <span id='jxnrgh'></span>
     
   <% 
   
   	Call drawSepLine() 		
 			'Call renderChart("FusionCharts/Column3D.swf?ChartNoDataText=点击饼图中的分类,显示每月分类工作量", "", "<chart></chart>", "dclassDetails", 400, 250, false, true)
  		  call  renderChart("FusionCharts/MSColumn3DLineDY.swf?ChartNoDataText=点击饼图中的分类,显示每月分类工作量", "","<chart></chart>","dclassDetails", 500, 325, false, true)
 
	%>
      
	</td></tr>
 
    
    
    
    
    
    
    
    
    
    
    
		
	 <%	
end sub













sub totalsscj()
  
  
  
  
  %>
  
		<Div class='x-layout-panel' style='WIDTH: 100%;'>
		<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
		<tr class="x-grid-header">
		     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>
		    </tr>




  <tr class='x-grid-row x-grid-row-alt' >  <td  bgcolor="#EDF9D5"  align="center"  >
    
    <%
	'显示车间检修量
				
				strsscjXML  = "<chart caption='" &  intYear &  "年 车间检修比例'  subcaption='点击饼图中的车间,显示每月车间工作量'  palette='" & getPalette()&  "' animation='" &  getAnimationState()&  "'  showValues='0' numberPrefix='' formatNumberScale='0' showPercentInToolTip='0'>"
				strsscjXML =  strsscjXML&getsscjEmpXML(intYear,-1,false,true,false)
				strsscjXML =  strsscjXML&"<styles><definition><style type='font' name='CaptionFont' color='"&  getCaptionFontColor()&  "' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
				strsscjXML = strsscjXML& "</chart>"
				call renderChart("FusionCharts/Pie3D.swf", "",strsscjXML, "Topsscj", 500, 325, false, false)
				%>
    </td>
  </tr>
 
 
			<tr class='x-grid-row x-grid-row-alt' >



    <td align="center"  bgcolor="#EDF9D5" > <span id='jxnrgh'></span>
<%	
	
   	Call drawSepLine() 		
  		  call  renderChart("FusionCharts/MSColumn3DLineDY.swf?ChartNoDataText=点击饼图中的车间,显示每月车间的工作量", "","<chart></chart>","sscjDetails", 500, 325, false, true)
 
	%>
      
      
      </td></tr>
      
		</table><%
end sub	










sub totalssbz()
  
  
			
			

  
  %>
  
		<Div class='x-layout-panel' style='WIDTH: 100%;'>
	
    
    
    
    
    	<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1">
		<tr class="x-grid-header">
		     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>
		    </tr>




  <tr class='x-grid-row x-grid-row-alt' >  <td  bgcolor="#EDF9D5"  align="center"  >
    
    <%
			
			dwt.out "<div style='text-align:left'>维修一:"
	
    strSQL = "SELECT id,bzname FROM  bzname  where sscj=1 order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		bzname1=Replace(ors("bzname")," ","")
		dwt.out  "<a href='?action=totalssbz&sscj=1&ssbz="&ors("id")&"&Year="&Request("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
		oRs.MoveNext()
	Wend	
	
	
	
	
	
			dwt.out "<br>维修二:"
    strSQL = "SELECT id,bzname FROM  bzname  where sscj=2 order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		bzname1=Replace(ors("bzname")," ","")
		dwt.out  "<a href='?action=totalssbz&sscj=2&ssbz="&ors("id")&"&Year="&Request("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
		oRs.MoveNext()
	Wend	
	
		
		
		
		
			dwt.out "<br>维修三:"

    strSQL = "SELECT id,bzname FROM  bzname  where sscj=3 order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		bzname1=Replace(ors("bzname")," ","")
		dwt.out  "<a href='?action=totalssbz&sscj=3&ssbz="&ors("id")&"&Year="&Request("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
		oRs.MoveNext()
	Wend	
	




			dwt.out "<br>维修四:"



    strSQL = "SELECT id,bzname FROM  bzname  where sscj=4 order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		bzname1=Replace(ors("bzname")," ","")
		dwt.out  "<a href='?action=totalssbz&sscj=4&ssbz="&ors("id")&"&Year="&Request("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
		oRs.MoveNext()
	Wend	
	


			dwt.out "<br>综合:"



    strSQL = "SELECT id,bzname FROM  bzname  where sscj=5 order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		bzname1=Replace(ors("bzname")," ","")
		dwt.out  "<a href='?action=totalssbz&sscj=5&ssbz="&ors("id")&"&Year="&Request("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
		oRs.MoveNext()
	Wend	
	



			'dwt.out "<br>计量:"


    'strSQL = "SELECT id,bzname FROM  bzname  where sscj=6 order by bzname asc"	
	'Set oRs = Server.CreateObject("ADODB.Recordset")
	'oRs.Open strSQL, Conn
	
	'While not oRs.EOF				
	'	bzname1=Replace(ors("bzname")," ","")
	'	dwt.out  "<a href='?action=totalssbz&sscj=6&ssbz="&ors("id")&"&Year="&Request.Form("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
	'	oRs.MoveNext()
	'Wend	
	

dwt.out "</div>"

	'显示车间检修量
				
				
				
				sscj=request("sscj")
				if sscj="" then sscj=1
				ssbz=request("ssbz")
				if ssbz="" then
				  ssbz=10
				  bzname="供水班"
				  end if 
				
				
    strSQL = "SELECT bzname FROM  bzname  where id="&ssbz&" order by bzname asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
		bzname=Replace(ors("bzname")," ","")
		'dwt.out  "<a href='?action=totalssbz&ssbz="&ors("id")&"&Year="&Request.Form("Year")&"'>"&bzname1&"</a>&nbsp;&nbsp;"
				
				if sscj=1 then sscjname="维修一"
				if sscj=2 then sscjname="维修二"
				if sscj=3 then sscjname="维修三"
				if sscj=4 then sscjname="维修四"
				if sscj=5 then sscjname="综合"
				if sscj=6 then sscjname="计量"
				




				%>
    </td>
  </tr>
 
 
			<tr class='x-grid-row x-grid-row-alt' >



    <td align="center"  bgcolor="#EDF9D5" > <span id='jxnrgh'></span>
    
    
	<%	dwt.out "<br><div style='line-height:20px'><strong><font size=16>"&sscjname & "&nbsp;&nbsp;"& bzname&"</font></strong></div><br>"
    strSQL = "SELECT userid.username1 FROM  userid  where userid.levelzclass="&ssbz&" and levelzclass>0 order by username1 asc"	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	While not oRs.EOF				
		'Append the data
		
		username1=Replace(ors("username1")," ","")
		dwt.out  username1
		
	i=i+1
		
		
'		jxcountsql="SELECT count(jx_id) as countjxid FROM sbjx INNER JOIN sb ON sbjx.sb_id=sb.sb_id WHERE 1=1 And (((sb.sb_sscj)=1)) And (sbjx.jx_ren Like '%"&username1&"%' Or sbjx.jx_fzren Like '%"&username1&"%') and year(jx_date)=2013"
'		set rsjx=server.createobject("adodb.recordset")
'		rsjx.open jxcountsql,conn,1,1
'		dwt.out username1 & rsjx("countjxid")& "<br>"

	'显示月检修量
	strmonthXML  = "<chart caption='"&intYear&"年  " &username1& " 每月检修量'  palette='" & getPalette() & "' animation='" & getAnimationState()& "' subcaption='' formatNumberScale='0' numberPrefix='' showValues='0' seriesNameInToolTip='0'>"
	strmonthXML= strmonthXML& getYgBymonth(intYear,username1)
	strmonthXML =strmonthXML& "<styles><definition><style type='font' color='" & getCaptionFontColor() & "' name='CaptionFont' size='15' /><style type='font' name='SubCaptionFont' bold='0' /></definition><application><apply toObject='caption' styles='CaptionFont' /><apply toObject='SubCaption' styles='SubCaptionFont' /></application></styles>"
	strmonthXML =strmonthXML &"</chart>"
	call  renderChart("FusionCharts/MSColumn3DLineDY.swf", "",strmonthXML,"SalesBymonth"&i, 500, 325, false, true)
	
	dwt.out "<br><br>"
		oRs.MoveNext()
	Wend	
	
	%>
      
      
      </td></tr>
      
		</table><%
end sub	
%>



</body></html>





