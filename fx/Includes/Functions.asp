<%
'This page contains functions which would be used by other pages
'We can store application states, messages and constants here
'Or create functions to handle data.

'escapeXML function helps you escape special characters in XML
Function escapeXML(strItem, forDataURL)
	'Convert ' to &apos; if dataURL
	if forDataURL=true then
		strItem = Replace(strItem,"'","&apos;")	
	else
		'Else for dataXML 		
		'Convert % to %25
		strItem = Replace(strItem,"%","%25")	
		'Convert ' to %26apos;
		strItem = Replace(strItem,"'","%26apos;")
		'Convert & to %26
		strItem = Replace(strItem,"&","%26")
	end if
	'Common replacements
	strItem = Replace(strItem,"<","&lt;")
	strItem = Replace(strItem,">","&gt;")
	'We've not considered any special characters here. 
	'You can add them as per your language and requirements.
	'Return
	escapeXML = strItem
End Function

'getPalette method returns a value between 1-5 depending on which
'paletter the user wants to plot the chart with. 
'Here, we just read from Session variable and show it
'In your application, you could read this configuration from your 
'User Configuration Manager, database, or global application settings
Function getPalette()
	Dim palette
	If Session("palette")="" then
		palette = "2"
	else
		palette = Session("palette")
	end if
	'Return
	getPalette = palette
End Function

'getAnimationState returns 0 or 1, depending on whether we've to
'animate chart. Here, we just read from Session variable and show it
'In your application, you could read this configuration from your 
'User Configuration Manager, database, or global application settings
Function getAnimationState()
	Dim animation
	If Session("animation")<>"0" then
		animation = "1"
	else
		animation = "0"
	end if
	'Return	
	getAnimationState = animation
End Function

'getCaptionFontColor function returns a color code for caption. Basic
'idea to use this is to demonstrate how to centralize your cosmetic 
'attributes for the chart
Function getCaptionFontColor()
	'Return a hex color code without #
	getCaptionFontColor = "666666"
	'FFC30C - Yellow Color
End Function
%>