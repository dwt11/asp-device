<%
DIM conn,connstr,db
  '获得当前日期的前一天，并格式化输出YYYYMMDD）   
  'bdate=DATE () - 1
  bdate=request("date")
  'response.Write bdate
  Dim   sTime   
  Dim   mTime   
  Dim   dTime   
  Dim   rTime   
  sTime=FormatDateTime(bdate,   vbShortDate)   
  mTime=month(sTime)   
  dTime=day(sTime)  
  IF   mTime<10   THEN   
        mTime="0"&mTime   
  End   IF   
  IF   dTime<10   THEN   
        dTime="0"&dTime   
  End   IF  
  
   
'nowmonth=year(sTime)'&mTime
nowmonth=2010
mTime="06"
dTime=16


db=nowmonth  '查询报表的月份
nowday=year(sTime)&mTime&dTime     '查询报表的天
nowday1=  year(sTime)&"-"&mTime&"-"&dTime
'response.Write nowday1
'判断查询天的数据库文件是否存在，如果存在输出报表，供服务器获取
datafolder=db&"/"&mTime&dTime

response.write datafolder

folderPath=Server.MapPath(datafolder)  
datafile=bianli(Server.MapPath("/")+"/"&datafolder)'调用bianli()函数,遍历对应日期目录下的DBF文件
		'response.Write datafile&"<br>"
	data2=Split(datafile,".dbf")
	

'dataday=GetBody(datafile,"<date>","</date>",false,false) 


'Set fso   =   CreateObject("Scripting.FileSystemObject")   
'if fso.FileExists(folderPath)   then
'	'call outdata(db,nowday) 
'else
'	response.Write bdate&"报表在轨道衡处未生成"&db&"/"&NOWDAY
'end   if   
'Set   fso   =   nothing

if datafile<>""   then
		response.Write "<day>"&nowday1&"</day>"
		response.Write "<data>"
	For i =0  To UBound(data2)-1
		'Set fso   =   CreateObject("Scripting.FileSystemObject")   
		'if fso.FileExists(folderPath)   then
		call outdata(datafolder,data2(i)) 
		'response.Write data2(i)
			'call outdata("2010/0607",rtrim(ltrim(data2(i)))) 
		response.Write data2(i)&"/"&i&"/"
		'end if  
	next
		response.Write "</data>"
		'	response.Write data2(0)
else	

	response.Write bdate&"报表在轨道衡处未生成"
	
end   if   
function outdata(db,nowday)
		set rs=nothing
Set conn = Server.CreateObject("ADODB.Connection")
	connstr = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF; SourceDB=" & Server.MapPath(db) &";Exclusive=No"
	conn.Open connstr
	set rs=server.createobject("adodb.recordset")
	 
		 
	
	
	SQL="select * from '"&nowday&".dbf'"
	rs.open SQL,conn,0,3
	if rs.eof or rs.bof  then 
	else
		do while not rs.eof or rs.bof 
			ii=ii+1
			response.Write rs("Xuhao")&","
			response.Write rs("chehao")&","
			response.Write rs("danwei")&","
			response.Write rs("wupin")&ii&","
			response.Write rs("sudu")&","
			response.Write rs("maozhs")&","
			response.Write rs("pizhs")&","
			response.Write rs("zaizhs")&","
			response.Write rs("jingzhs")&","
			response.Write rs("lusun")&","
			response.Write rs("yingkui")&","	
			response.Write rs("xch")&","
			response.Write rs("nian")&"-"
			response.Write rs("yue")&"-"	
			response.Write rs("ri")&" "	
			response.Write rs("shi")&":"
			response.Write rs("fen")
response.Write "$"
			rs.movenext
			'if  not rs.eof then 
		loop
	end if 
	rs.close
	set rs=nothing
end function





 
function bianli(path)
        dim fso            'fso对象
        dim objFolder      '文件夹对象
        dim objFiles       '文件集合
        dim objFile        '文件对象
		on error resume next

        set fso=server.CreateObject("Scripting.FileSystemObject")    
        set objFolder=fso.GetFolder(path)'创建文件夹对象
		 'Response.Write ("<swf>")
				set objFiles=objFolder.Files
				for each objFile in objFiles
					bianli=bianli&objFile.name
				next
	  ' Response.Write ("</swf>")
        set objFolder=nothing
        set objSubFolders=nothing
        set fso=nothing
end function




%>