<%
'年检修量
Function getSalesByYear()
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT Year(jx_date) As Year1, count(jx_id) As Total, count(jx_id) as Quantity FROM sbjx GROUP BY Year(jx_date)  ORDER BY Year(jx_date)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname=''>"
	strQtyDS = "<dataset seriesName='' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & ors("Year1") & "'/>"		
		
		'Generate the link
		'strLink = Server.URLEncode("javaScript:updateCharts(" & ors("Year1") & ");")
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesByYear = strXML
End Function






'每月检修量
Function getSalesBymonth(intYear)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT month(jx_date) As month1, count(jx_id) As Total, count(jx_id) as Quantity FROM sbjx where year(jx_date)="&intyear&" GROUP BY month(jx_date)  ORDER BY month(jx_date)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname=''>"
	strQtyDS = "<dataset seriesName='' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & ors("month1") & "'/>"		
		
		'Generate the link
		'strLink = Server.URLEncode("javaScript:updateCharts(" & ors("Year1") & ");")
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesBymonth = strXML
End Function











'某年中的分类检修比例
Function getSalePerEmpXML(intYear, howMany, slicePies, addJSLinks,forDataURL)
	strSQL = "SELECT count(jx_id) As Total FROM sbjx where Year(jx_date)="&intYear
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	
		jxtotal =  ors("Total")   '一年检修总量		

	
	Dim oRs, strSQL
	Dim strXML, count
	count = 0
	if howMany=-1 then
		strSQL = "SELECT sbclass_id,sbclass_name ,count(jx_id) As Total FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)=" & intYear & " group by sbclass_name,sbclass_id order by count(jx_id) desc"	
	else
		strSQL = "SELECT TOP " & howMany & " sbclass_name ,count(jx_id) As Total FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)=" & intYear & " group by sbclass_name order by count(jx_id) desc"	
	end if
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	Dim strLink
	
	Dim slicedOut
	While not oRs.EOF				
		'Append the data
		count = count+ 1
		
		'If link is to be added
		if addJSLinks=true then
			strLink = " link='javascript:updateChartClass(" & oRs("sbclass_id") & "," &intYear & ");' "
			'strLink = " link='javascript:updateChartClass(" & oRs("sbclass_id") & "," &intYear & ");' "
		else
			strLink = ""
		end if
		'If top 2 employees, then sliced out				
		if slicePies and count<3 then 
			slicedOut="1" 			
		else
			slicedOut="0"
		end if
		bl=FormatNumber(cInt(ors("Total"))/cint(jxtotal)*100,2,-1,0,0)
		strXML = strXML & "<set label='" & escapeXML(ors("sbclass_name"),forDataURL) & "&nbsp;"&bl&"%' value='" & Int(ors("Total")) & "' isSliced='" & slicedOut & "' " & strLink & " />"
		'strXML = strXML & "<set label='" & escapeXML("",forDataURL) & "&nbsp;"&bl&"%' value='" & Int(ors("Total")) & "' isSliced='" & slicedOut & "' " & strLink & " />"
		oRs.MoveNext()
	Wend	
	getSalePerEmpXML = strXML
End Function



'每月某分类检修量
Function getclassBymonth(intYear,sbdclass)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT month(jx_date) As month1, count(jx_id) As Total,count(jx_id) As Quantity FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)=" & intYear & "  and sb.sb_dclass="&sbdclass &"  GROUP BY month(jx_date)  ORDER BY month(jx_date)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname=''>"
	strQtyDS = "<dataset seriesName='' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & ors("month1") & "'/>"		
		
		'Generate the link
		'strLink = Server.URLEncode("javaScript:updateCharts(" & ors("Year1") & ");")
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getclassBymonth = strXML
End Function




'已经不用了某年中的每月分类检修量 top5
Function getEmployeeBymonth(intYear, howMany)
	



	Dim oRs, strSQL
	Dim strXML
	Dim strCat
	Dim strAmtDS, strQtyDS
	
	strCat = "<categories>"
	
	
	
	if howMany=-1 then
		strSQL = "SELECT  sb.sb_dclass,sbclass.sbclass_name FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)= " & intYear & " group by sb.sb_dclass,sbclass_name order by count(jx_id) desc"	
	else
		strSQL = "SELECT TOP " & howMany & " sb.sb_dclass,sbclass.sbclass_name FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)= " & intYear & " group by sb.sb_dclass,sbclass_name order by count(jx_id) desc"	
	end if
	Set oRs = Server.CreateObject("ADODB.Recordset")
	'response.Write strsql
	oRs.Open strSQL, Conn
	
	Dim strLink
	
	Dim slicedOut
	While not oRs.EOF				
    	strAmtDS = strAmtDS&"<dataset seriesname='"&ors("sbclass_name")&"'>"
		for i=1 to 12
			  strSQL1 = "SELECT  count(jx_id) As Total FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN sbclass ON sb.sb_dclass = sbclass.sbclass_id where year(jx_date)=" & intYear & " and month(jx_date)="&i&" and sb.sb_dclass="&ors("sb_dclass")
			  Set oRs1 = Server.CreateObject("ADODB.Recordset")
			  oRs1.Open strSQL1, Conn
		     strAmtDS = strAmtDS & "<set value='"&ors1("total")&"' />"	
	
			  oRs1.Close()
			  Set oRs1 = nothing
			 	
		next 
	strAmtDS = strAmtDS & "</dataset>"
		oRs.MoveNext()
	Wend		
	
	
	
		for i=1 to 12
			strCat = strCat & "<category label='" & i & "月'/>"		
		next 
	
	
	
	strCat = strCat & "</categories>"
	strXML = strCat & strAmtDS 
	
	oRs.Close()
	Set oRs = nothing
	
	getEmployeeBymonth = strXML



	'getEmployeeBymonth = "<categories><category label='1' /><category label='2' /><category label='3' /><category label='4' /><category label='5' /><category label='6' /><category label='7' /><category label='8' /><category label='9' /><category label='10' /><category label='11' /><category label='12' /></categories><dataset seriesName='2012收入'><set value='700' /><set value='5498' /><set value='14100' /><set value='24441' /><set value='31571' /><set value='3160' /><set value='0' /><set value='0' /><set value='395' /><set value='910' /><set value='3051' /><set value='2550' /></dataset><dataset seriesName='支出'><set value='0' /><set value='0' /><set value='1445' /><set value='1084' /><set value='1520' /><set value='200' /><set value='301' /><set value='1500' /><set value='365' /><set value='0' /><set value='0' /><set value='250' /></dataset><dataset seriesName='2011收入' parentYAxis='S'><set value='480'/><set value='706'/><set value='9653'/><set value='21021'/><set value='28870'/><set value='4030'/><set value='0'/><set value='260'/><set value='700'/><set value='1900'/><set value='3500'/><set value='3870'/></dataset>"



End Function









'某年中的车间检修比例
Function getsscjEmpXML(intYear, howMany, slicePies, addJSLinks,forDataURL)
	
	
	
	
	
	
	strSQL = "SELECT count(jx_id) As Total FROM sbjx where Year(jx_date)="&intYear
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	
		jxtotal =  ors("Total") 		

	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store entire XML Data
	Dim strXML, count
	count = 0
	'Retrieve the data
	if howMany=-1 then
		strSQL = "SELECT levelid,levelname ,count(jx_id) As Total FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN levelname ON sb.sb_sscj = levelname.levelid where year(jx_date)=" & intYear & " group by levelname,levelid order by count(jx_id) desc"	
	else
		strSQL = "SELECT TOP " & howMany & " SELECT levelname ,count(jx_id) As Total FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN levelname ON sb.sb_sscj = levelname.levelid where year(jx_date)=" & intYear & " group by levelname,levelid order by count(jx_id) desc"	
	end if
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn
	
	'Link to be added
	Dim strLink
	
	'Whether sliced
	Dim slicedOut
	'Create the XML data document containing only data
	'We add the <chart> element in the calling function, depending on needs.	
	While not oRs.EOF				
		'Append the data
		count = count+ 1
		
		'If link is to be added
		if addJSLinks=true then
			strLink = " link='javascript:updateChartSscj(" & oRs("levelid") & "," &intYear & ");' "
		else
			'strLink = ""
		end if
		'If top 2 employees, then sliced out				
		if slicePies and count<3 then 
			slicedOut="1" 			
		else
			slicedOut="0"
		end if
		bl=FormatNumber(cInt(ors("Total"))/cint(jxtotal)*100,2,-1,0,0)
		strXML = strXML & "<set label='" & escapeXML(ors("levelname"),forDataURL) & "&nbsp;"&bl&"%' value='" & Int(ors("Total")) & "' isSliced='" & slicedOut & "' " & strLink & " />"
		oRs.MoveNext()
	Wend	
	getsscjEmpXML = strXML
End Function








'每月某车间检修量
Function getsscjBymonth(intYear,sscj)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS

	strSQL = "SELECT month(jx_date) As month1, count(jx_id) As Total,count(jx_id) As Quantity FROM ( sbjx INNER JOIN sb ON sbjx.sb_id = sb.sb_id ) INNER JOIN levelname ON sb.sb_sscj = levelname.levelid where year(jx_date)=" & intYear & "  and levelname.levelid="&sscj &"  GROUP BY month(jx_date)  ORDER BY month(jx_date)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname=''>"
	strQtyDS = "<dataset seriesName='' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & ors("month1") & "'/>"		
		
		'Generate the link
		'strLink = Server.URLEncode("javaScript:updateCharts(" & ors("Year1") & ");")
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getsscjBymonth = strXML
End Function








'员工每月检修量
Function getYgBymonth(intYear,username1)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT month(jx_date) As month1, count(jx_id) As Total, count(jx_id) as Quantity FROM sbjx  WHERE  (sbjx.jx_ren Like '%"&username1&"%' Or sbjx.jx_fzren Like '%"&username1&"%') and year(jx_date)="&intyear&" GROUP BY month(jx_date)  ORDER BY month(jx_date)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, Conn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname=''>"
	strQtyDS = "<dataset seriesName='' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & ors("month1") & "'/>"		
		
		'Generate the link
		'strLink = Server.URLEncode("javaScript:updateCharts(" & ors("Year1") & ");")
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getYgBymonth = strXML
End Function


















'getSalesByCountryXML function returns the XML Data for sales
'for a given country in a given year.
Function getSalesByCountryXML(intYear, howMany, addJSLinks,forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	if howMany=-1 then
		strSQL = "SELECT  c.country, round(SUM(d.quantity*p.UnitPrice*(1-d.discount)),0) As Total, SUM(d.quantity) as Quantity FROM FC_Customers as c, FC_Products as p, FC_Orders as o, FC_OrderDetails as d WHERE YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and c.customerid= o.customerid and o.orderid= d.orderid GROUP BY c.country ORDER BY SUM(d.quantity*p.UnitPrice*(1- d.discount)) DESC"	
	else
		strSQL = "SELECT TOP " & howMany & " c.country, round(SUM(d.quantity*p.UnitPrice*(1-d.discount)),0) As Total, SUM(d.quantity) as Quantity FROM FC_Customers as c, FC_Products as p, FC_Orders as o, FC_OrderDetails as d WHERE YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and c.customerid= o.customerid and o.orderid= d.orderid GROUP BY c.country ORDER BY SUM(d.quantity*p.UnitPrice*(1- d.discount)) DESC"	
	end if
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Amount'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & escapeXML(ors("Country"),forDataURL) & "'/>"		
		
		'If JavaScript links are to be added
		if addJSLinks=true then			
			'Generate the link
			'TRICKY: We're having to escape the " character using chr(34) character.
			'In HTML, the data is provided as chart.setXMLData(" - so " is already used and un-terminated
			'For each XML attribute, we use '. So ' is used in <set link='
			'Now, we've to pass Country Name to JavaScript function, so we've to use chr(34)
			strLink = Server.URLEncode("javaScript:updateChart(" & intYear & "," & chr(34) & ors("Country") &  chr(34) & ");")
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"
		else
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"
		end if		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesByCountryXML = strXML
End Function

'getSalesByCountryCityXML function generates the XML data for sales
'by city within the given country, for the given year.
Function getSalesByCountryCityXML(intYear, country,forDataURL)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT  c.city, round(SUM(d.quantity*p.UnitPrice*(1-d.discount)),0) As Total, SUM(d.quantity) as Quantity  FROM FC_Customers as c, FC_Products as p, FC_Orders as o, FC_OrderDetails as d WHERE YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and c.customerid= o.customerid and o.orderid= d.orderid and c.country='" & country & "' GROUP BY c.city ORDER BY SUM(d.quantity*p.UnitPrice*(1- d.discount)) DESC"		
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Amount'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & escapeXML(ors("City"),forDataURL) & "'/>"		
		
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesByCountryCityXML = strXML
End Function

'getSalesByCountryCustomerXML function generates the XML data for sales
'by customers within the given country, for the given year.
Function getSalesByCountryCustomerXML(intYear, country, forDataURL)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT c.companyName as CustomerName, SUM(d.quantity*p.UnitPrice) As Total, SUM(d.Quantity) As Quantity FROM FC_Customers as c, FC_OrderDetails as d, FC_Orders as o, FC_products as p WHERE YEAR(OrderDate)=" & intYear & " and c.customerid=o.customerid and o.orderid=d.orderid and d.productid=p.productid and c.country='" & country & "' GROUP BY c.CompanyName ORDER BY SUM(d.quantity*p.UnitPrice) DESC"		
		
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Amount'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		'Since customers name are long, we truncate them to 5 characters and then show ellipse
		'The full name is then shown as toolText
		strCat = strCat & "<category label='" & escapeXML(Left(ors("CustomerName"),5) & "...", forDataURL) & "' toolText='" & escapeXML(ors("CustomerName"),forDataURL) & "'/>"				
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesByCountryCustomerXML = strXML
End Function

'getExpensiveProdXML method returns the 10 most expensive products
'in the database along with the sales quantity of those products
'for the given year
Function getExpensiveProdXML(intYear, howMany, forDataURL)
'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT TOP " & howMany & " p.ProductName, p.UnitPrice, SUM(d.quantity) as Quantity FROM FC_Products p, FC_Orders as o, FC_OrderDetails d where YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and o.orderid= d.orderid group by p.productname,p.UnitPrice  order by p.UnitPrice desc"
		
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Unit Price'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & escapeXML(ors("ProductName"),forDataURL) & "'/>"		
		
		strAmtDS = strAmtDS & "<set value='" & ors("UnitPrice") & "' />"
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getExpensiveProdXML = strXML
End Function

'getInventoryByCatXML function returns the inventory of all items
'and their respective quantity
Function getInventoryByCatXML(addJSLinks,forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "select  c.categoryname,round((sum(p.UnitsInStock)),0) as Quantity, round((sum(p.UnitsInStock*p.UnitPrice)),0) as Total from FC_categories as c , FC_products as p where c.categoryid=p.categoryid group by c.categoryname order by (sum(p.UnitsInStock*p.UnitPrice)) Desc"
		
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Cost of Inventory'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & escapeXML(ors("CategoryName"),forDataURL) & "'/>"		
		
		'If JavaScript links are to be added
		if addJSLinks=true then			
			'Generate the link
			'TRICKY: We're having to escape the " character using chr(34) character.
			'In HTML, the data is provided as chart.setXMLData(" - so " is already used and un-terminated
			'For each XML attribute, we use '. So ' is used in <set link='
			'Now, we've to pass Country Name to JavaScript function, so we've to use chr(34)
			strLink = Server.URLEncode("javaScript:updateChart("& chr(34) & ors("CategoryName") &  chr(34) & ");")
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"
		else
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"
		end if		
		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getInventoryByCatXML = strXML
End Function

'getInventoryByProdXML function returns the inventory of all items
'within a given category and their respective quantity
Function getInventoryByProdXML(catName,forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "select p.productname,round((sum(p.UnitsInStock)),0) as Quantity , round((sum(p.UnitsInStock*p.UnitPrice)),0) as Total from FC_Categories as c , FC_products as p where c.categoryid=p.categoryid and c.categoryname='" & catName & "' group by p.productname having sum(p.UnitsInStock)>0"	
		
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Cost of Inventory'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store short name	
	Dim shortName
	
	'Iterate through each data row
	While not oRs.EOF		
		'Product Names are long - so show 8 characters and ... and show full thing in tooltip
		if Len(ors("productname"))>8 then
			shortName = escapeXML(Left(ors("productname"),8) & "...",forDataURL)
		else
			shortName = escapeXML(ors("productname"),forDataURL)
		end if
		strCat = strCat & "<category label='" & shortName & "' toolText='" & escapeXML(ors("productname"),forDataURL) & "'/>"		
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getInventoryByProdXML = strXML
End Function

'getSalesByCityXML function returns the XML Data for sales
'for all cities in a given year.
Function getSalesByCityXML(intYear, howMany, forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	
	if howMany=-1 then
		strSQL = "SELECT c.city, SUM(d.quantity*p.UnitPrice) As Total FROM FC_Customers as c, FC_Products as p, FC_Orders as o, FC_OrderDetails as d   WHERE YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and c.customerid = o.customerid and o.orderid= d.orderid GROUP BY c.city order by SUM(d.quantity*p.UnitPrice) desc"	
	else
		strSQL = "SELECT top " & howMany & " c.city, SUM(d.quantity*p.UnitPrice) As Total FROM FC_Customers as c, FC_Products as p, FC_Orders as o, FC_OrderDetails as d   WHERE YEAR(OrderDate)=" & intYear & " and d.productid= p.productid and c.customerid = o.customerid and o.orderid= d.orderid GROUP BY c.city order by SUM(d.quantity*p.UnitPrice) desc"
	end if
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF		
		strXML = strXML & "<set label='" & escapeXML(ors("City"),forDataURL) & "' value='" & ors("Total") & "' />"		
		oRs.MoveNext()
	Wend

	oRs.Close()
	Set oRs = nothing
	
	getSalesByCityXML = strXML
End Function

'getYrlySalesByCatXML function returns the XML Data for sales
'for a given country in a given year.
Function getYrlySalesByCatXML(intYear, addJSLinks,forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT g.categoryId,g.categoryname,SUM(d.quantity*p.UnitPrice) as Total, SUM(d.quantity) As Quantity FROM FC_categories as g, FC_products as p, FC_orders as o, FC_orderdetails as d  WHERE YEAR(OrderDate)=" & intYear & " and d.productid=p.productid and g.categoryid=p.categoryid and o.orderid=d.orderid GROUP BY g.categoryId,g.categoryname ORDER BY SUM(d.quantity*p.UnitPrice) DESC"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Revenue'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store link
	Dim strLink
	
	'Iterate through each data row
	While not oRs.EOF
		strCat = strCat & "<category label='" & escapeXML(ors("CategoryName"),forDataURL) & "'/>"		
		
		'If JavaScript links are to be added
		if addJSLinks=true then			
			'Generate the link
			strLink = Server.URLEncode("javaScript:updateChart(" & intYear & "," & ors("CategoryId") & ");")
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' link='" & strLink & "'/>"
		else
			strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"
		end if		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getYrlySalesByCatXML = strXML
End Function

'getSalesByProdCatXML function returns the sales of all items
'within a given category in a year and their respective quantity
Function getSalesByProdCatXML(intYear,catID,forDataURL)
	'Initialize database objects
	Dim oRs, strSQL
	'Variable to store XML Data
	Dim strXML
	'To store categories
	Dim strCat
	'To store amount Dataset & quantity dataset
	Dim strAmtDS, strQtyDS
	
	strSQL = "SELECT g.categoryname,p.productname,round(sum(d.quantity),0) as quantity, round(SUM(d.quantity*p.UnitPrice),0) As Total FROM FC_Categories as g,  FC_Products as p, FC_Orders as o, FC_OrderDetails as d WHERE year(o.OrderDate)=" & intYear & " and g.CategoryID=" & catId & " and d.productid= p.productid and g.categoryid= p.categoryid and o.orderid=d.orderid GROUP BY g.categoryname,p.ProductName"
		
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn

	'Initialize <categories> element
	strCat = "<categories>"
	
	'Initialize datasets
	strAmtDS = "<dataset seriesname='Revenue'>"
	strQtyDS = "<dataset seriesName='Quantity' parentYAxis='S'>"
	
	'Variable to store short name	
	Dim shortName
	
	'Iterate through each data row
	While not oRs.EOF		
		'Product Names are long - so show 8 characters and ... and show full thing in tooltip
		if Len(ors("productname"))>8 then
			shortName = escapeXML(Left(ors("productname"),8) & "...",forDataURL)
		else
			shortName = escapeXML(ors("productname"),forDataURL)
		end if
		strCat = strCat & "<category label='" & shortName & "' toolText='" & escapeXML(ors("productname"),forDataURL) & "'/>"
		strAmtDS = strAmtDS & "<set value='" & ors("Total") & "' />"		
		strQtyDS = strQtyDS & "<set value='" & oRs("Quantity") & "'/>"
		oRs.MoveNext()
	Wend
	'Closing elements
	strCat = strCat & "</categories>"
	strAmtDS = strAmtDS & "</dataset>"
	strQtyDS = strQtyDS & "</dataset>"
	'Entire XML - concatenation
	strXML = strCat & strAmtDS & strQtyDS
	
	oRs.Close()
	Set oRs = nothing
	
	getSalesByProdCatXML = strXML
End Function


'getEmployeeName function returns the name of an employee based
'on his id.
Function getEmployeeName(empId)
	'Initialize database objects
	Dim oRs, strSQL	
	Dim name
		
	'Retrieve the data
	strSQL = "SELECT FirstName, lastname FROM FC_Employees where EmployeeID=" & empId 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn
	if not oRs.EOF then
		name = oRs("FirstName") & " " & oRs("LastName")
	else
		name = " N/A "
	end if
	Set oRs = nothing
	'Return
	getEmployeeName = name
End Function

'getCategoryName function returns the category name for a given category
'id
Function getCategoryName(catId)
	'Initialize database objects
	Dim oRs, strSQL	
	Dim name
		
	'Retrieve the data
	strSQL = "SELECT CategoryName FROM FC_Categories where CategoryId=" & catId
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open strSQL, oConn
	if not oRs.EOF then
		name = oRs("CategoryName")
	else
		name = " "
	end if
	Set oRs = nothing
	'Return
	getCategoryName = name
End Function
%>