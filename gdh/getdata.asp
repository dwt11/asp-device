<%
'getdatA "2009-6-1"
function getdatA(dateinput)
'	set rs=server.createobject("adodb.recordset")
'	sql="select * from issave where issave and day1=#"&formatDateTime(dateinput,vbShortDate) &"#"
'	rs.open sql,connjlhs,1,1
'	if rs.eof and rs.bof then 
		'body= gethttppage("http://172.16.10.131/index55.asp?date="&dateinput)
		body= gethttppage("http://172.16.10.131/index55.asp?date="&dateinput)
		'response.Write body&"<br>"
		dataday=GetBody(body,"<day>","</day>",false,false)   '��ȡ��������
		if  dataday<>false  then 
			call savedate(dataday,body)
		else
			response.Write "����⴦���ɵı����ʽ������"	
		end if 
'	else
'		'response.Write "�ѱ����"
'	end if 
'	'response.Write "<br>�������ݳɹ�"	
end function

sub savedate(dataday,body)
	data=GetBody(body,"<data>","</data>","","") 
	data1=Split(data,"$")
	localdatanumb=connjlhs.Execute("SELECT count(id) FROM qch where day=#"&formatDateTime(dataday,vbShortDate) &"#")(0)
    remotedatanumb=UBound(data1)
	'response.Write localdatanumb&"------"&UBound(data1)
	
	if localdatanumb<remotedatanumb then 
		if localdatanumb=0 then befornumb=LBound(data1)
		if localdatanumb<>0 then befornumb=localdatanumb
	   ' response.Write "------"&befornumb
		For i =befornumb  To UBound(data1)-1
		   'Response.Write("����ֵ�����е�Ԫ��[" & i & "]��" & data1(i)& "<br>")
		   data2=Split(data1(i),",")
'			For ii = LBound(data2) To UBound(data2)
'			   Response.Write("<br>����ֵ�����е�Ԫ��[" & ii & "]��" & data2(ii)& "<br>")
'			Next
			savedata dataday,data2(1),data2(2),data2(3),data2(4),data2(5),data2(6),data2(7),data2(8),data2(9),data2(10),data2(11),data2(12)
		Next
		set rsadd=server.createobject("adodb.recordset")
		sqladd="select * from issave" 
		rsadd.open sqladd,connjlhs,1,3
		rsadd.addnew
		rsadd("day1")=dataday
		rsadd("issave")=true
		rsadd.update
		rsadd.close
		set rsadd=nothing
	end if 	
	'response.write dataday &"<br>"
	'response.write data &"<br>"
end sub



function savedata(dataday,chehao,danwei,wupin,sudu,maozhs,pizhs,zaizhs,jingzhs,lusun,yingkui,xch,gbdate)    
'		db="jlhs.mdb"
'		Set connjlhs = Server.CreateObject("ADODB.Connection")
'		Conns = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
'		connjlhs.Open conns
      set rsadd=server.createobject("adodb.recordset")
      sqladd="select * from qch" 
      rsadd.open sqladd,connjlhs,1,3
      rsadd.addnew
		on error resume next
      rsadd("day")=dataday
      rsadd("chehao")=chehao
      rsadd("danwei")=danwei
      rsadd("wupin")=wupin
      rsadd("sudu")=sudu
      rsadd("maozhs")=maozhs
      rsadd("pizhs")=pizhs
      rsadd("zaizhs")=zaizhs
      rsadd("jingzhs")=jingzhs
	  rsadd("lusun")=lusun
      rsadd("yingkui")=yingkui
      rsadd("xch")=xch
      rsadd("gbdate")=gbdate
      
      rsadd.update
      rsadd.close
	  

     ' set rsadd=nothing
      
	  'dwt.savesl "��������","Զ�̲ɼ����ݲ�����",dateday
end function







			








'==================================================
'��������GetHttpPage
'�� �ã���ȡ��ҳԴ��
'�� ����HttpUrl ------��ҳ��ַ
'==================================================
Function GetHttpPage(HttpUrl)
	If IsNull(HttpUrl)=True Or Len(HttpUrl)<18 Or HttpUrl="$False$" Then
		GetHttpPage="$False$"
		Exit Function
	End If
	Dim Http
	Set Http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",HttpUrl,False
	Http.Send()
	If Http.Readystate<>4 then
		Set Http=Nothing 
		GetHttpPage="$False$"
		Exit function
	End if
	GetHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
	Set Http=Nothing
	If Err.number<>0 then
		Err.Clear
	End If
End Function


'==================================================
'��������BytesToBstr
'�� �ã�����ȡ��Դ��ת��Ϊ����
'�� ����Body ------Ҫת���ı���
'�� ����Cset ------Ҫת��������
'==================================================
Function BytesToBstr(Body,Cset)
	Dim Objstream
	Set Objstream = Server.CreateObject("adodb.stream")
	objstream.Type = 1
	objstream.Mode =3
	objstream.Open
	objstream.Write body
	objstream.Position = 0
	objstream.Type = 2
	objstream.Charset = Cset
	BytesToBstr = objstream.ReadText 
	objstream.Close
	set objstream = nothing
End Function
'==================================================
'��������UrlEncoding
'�� �ã�ת������
'==================================================
Function UrlEncoding(DataStr)
	Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
	StrReturn = ""
	For Si = 1 To Len(DataStr)
		ThisChr = Mid(DataStr,Si,1)
		If Abs(Asc(ThisChr)) < &HFF Then
			StrReturn = StrReturn & ThisChr
		Else
			InnerCode = Asc(ThisChr)
			If InnerCode < 0 Then
				InnerCode = InnerCode + &H10000
			End If
			Hight8 = (InnerCode And &HFF00)\ &HFF
			Low8 = InnerCode And &HFF
			StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
		End If
	Next
	UrlEncoding = StrReturn
End Function

'==================================================
'��������GetBody
'�� �ã���ȡ�ַ���
'�� ����ConStr ------��Ҫ��ȡ���ַ���
'�� ����StartStr ------��ʼ�ַ���
'�� ����OverStr ------�����ַ���
'�� ����IncluL ------�Ƿ����StartStr
'�� ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
	If ConStr=False or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
		GetBody=False
		Exit Function
	End If
	Dim ConStrTemp
	Dim Start,Over
	ConStrTemp=Lcase(ConStr)
	StartStr=Lcase(StartStr)
	OverStr=Lcase(OverStr)
	Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
	If Start<=0 then
		GetBody=False
		Exit Function
	Else
		If IncluL=False Then
			Start=Start+LenB(StartStr)
		End If
	End If
	Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
	If Over<=0 Or Over<=Start then
		GetBody=False
		Exit Function
	Else
		If IncluR=True Then
			Over=Over+LenB(OverStr)
		End If
	End If
	GetBody=MidB(ConStr,Start,Over-Start)
End Function


'==================================================
'��������GetArray
'�� �ã���ȡ���ӵ�ַ����$Array$�ָ�
'�� ����ConStr ------��ȡ��ַ��ԭ�ַ�
'�� ����StartStr ------��ʼ�ַ���
'�� ����OverStr ------�����ַ���
'�� ����IncluL ------�Ƿ����StartStr
'�� ����IncluR ------�Ƿ����OverStr
'==================================================
Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or IsNull(StartStr)=True Or IsNull(OverStr)=True Then
GetArray="$False$"
Exit Function
End If
Dim TempStr,TempStr2,objRegExp,Matches,Match
TempStr=""
Set objRegExp = New Regexp 
objRegExp.IgnoreCase = True 
objRegExp.Global = True
objRegExp.Pattern = "("&StartStr&").+?("&OverStr&")"
Set Matches =objRegExp.Execute(ConStr) 
For Each Match in Matches
TempStr=TempStr & "$Array$" & Match.Value
Next 
Set Matches=nothing

If TempStr="" Then
GetArray="$False$"
Exit Function
End If
TempStr=Right(TempStr,Len(TempStr)-7)
If IncluL=False then
objRegExp.Pattern =StartStr
TempStr=objRegExp.Replace(TempStr,"")
End if
If IncluR=False then
objRegExp.Pattern =OverStr
TempStr=objRegExp.Replace(TempStr,"")
End if
Set objRegExp=nothing
Set Matches=nothing

TempStr=Replace(TempStr,"""","")
TempStr=Replace(TempStr,"'","")
TempStr=Replace(TempStr," ","")
TempStr=Replace(TempStr,"(","")
TempStr=Replace(TempStr,")","")

If TempStr="" then
GetArray="$False$"
Else
GetArray=TempStr
End if
End Function


'==================================================
'��������DefiniteUrl
'�� �ã�����Ե�ַת��Ϊ���Ե�ַ
'�� ����PrimitiveUrl ------Ҫת������Ե�ַ
'�� ����ConsultUrl ------��ǰ��ҳ��ַ
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" or ConsultUrl="$False$" Then
DefiniteUrl="$False$"
Exit Function
End If
If Left(Lcase(ConsultUrl),7)<>"http://" Then
ConsultUrl= "http://" & ConsultUrl
End If
ConsultUrl=Replace(ConsultUrl,"\","/")
ConsultUrl=Replace(ConsultUrl,"://",":\\")
PrimitiveUrl=Replace(PrimitiveUrl,"\","/")

If Right(ConsultUrl,1)<>"/" Then
If Instr(ConsultUrl,"/")>0 Then
If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then 
Else
ConsultUrl=ConsultUrl & "/"
End If
Else
ConsultUrl=ConsultUrl & "/"
End If
End If
ConArray=Split(ConsultUrl,"/")

If Left(LCase(PrimitiveUrl),7) = "http://" then
DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
ElseIf Left(PrimitiveUrl,1) = "/" Then
DefiniteUrl=ConArray(0) & PrimitiveUrl
ElseIf Left(PrimitiveUrl,2)="./" Then
PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-2)
If Right(ConsultUrl,1)="/" Then 
DefiniteUrl=ConsultUrl & PrimitiveUrl
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
End If
ElseIf Left(PrimitiveUrl,3)="../" then
Do While Left(PrimitiveUrl,3)="../"
PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
Pi=Pi+1
Loop 
For Ci=0 to (Ubound(ConArray)-1-Pi)
If DefiniteUrl<>"" Then
DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
Else
DefiniteUrl=ConArray(Ci)
End If
Next
DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
Else
If Instr(PrimitiveUrl,"/")>0 Then
PriArray=Split(PrimitiveUrl,"/")
If Instr(PriArray(0),".")>0 Then
If Right(PrimitiveUrl,1)="/" Then
DefiniteUrl="http:\\" & PrimitiveUrl
Else
If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
DefiniteUrl="http:\\" & PrimitiveUrl
Else
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
End If
End If 
Else
If Right(ConsultUrl,1)="/" Then 
DefiniteUrl=ConsultUrl & PrimitiveUrl
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
End If
End If
Else
If Instr(PrimitiveUrl,".")>0 Then
If Right(ConsultUrl,1)="/" Then
If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=ConsultUrl & PrimitiveUrl
End If
Else
If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
End If
End If
Else
If Right(ConsultUrl,1)="/" Then
DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
End If 
End If
End If
End If
If Left(DefiniteUrl,1)="/" then
DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
End if
If DefiniteUrl<>"" Then
DefiniteUrl=Replace(DefiniteUrl,"//","/")
DefiniteUrl=Replace(DefiniteUrl,":\\","://")
Else
DefiniteUrl="$False$"
End If
End Function

'==================================================
'��������ReplaceSaveRemoteFile
'�� �ã��滻������Զ��ͼƬ
'�� ����ConStr ------ Ҫ�滻���ַ���
'�� ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
'�� ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
If ConStr="$False$" or ConStr="" or strInstallDir="" or strChannelDir="" Then
ReplaceSaveRemoteFile=ConStr
Exit Function
End If
Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

Set Re = New Regexp 
Re.IgnoreCase = True 
Re.Global = True
Re.Pattern ="<img.+?[^\>]>"
Set Matches =Re.Execute(ConStr) 
For Each Match in Matches
If TempStr<>"" then 
TempStr=TempStr & "$Array$" & Match.Value
Else
TempStr=Match.Value
End if
Next
If TempStr<>"" Then
TempArray=Split(TempStr,"$Array$")
TempStr=""
For Tempi=0 To Ubound(TempArray)
Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
Set Matches =Re.Execute(TempArray(Tempi)) 
For Each Match in Matches
If TempStr<>"" then 
TempStr=TempStr & "$Array$" & Match.Value
Else
TempStr=Match.Value
End if
Next
Next
End if
If TempStr<>"" Then
Re.Pattern ="src\s*=\s*"
TempStr=Re.Replace(TempStr,"")
End If
Set Matches=nothing
Set Re=nothing
If TempStr="" or IsNull(TempStr)=True Then
ReplaceSaveRemoteFile=ConStr
Exit function
End if
TempStr=Replace(TempStr,"""","")
TempStr=Replace(TempStr,"'","")
TempStr=Replace(TempStr," ","")

Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
DtNow=Now()
If SaveTf=True then
'***********************************
SavePath= strChannelDir & "/" & year(DtNow) & right("0" & month(DtNow),2) & "/"
   response.write "����·����" & savepath & "<br>"
Arr_Path=Split(SavePath,"/")
PathTemp=""
For Tempi=0 To Ubound(Arr_Path)
If Tempi=0 Then
PathTemp=Arr_Path(0) & "/"
ElseIf Tempi=Ubound(Arr_Path) Then
Exit For
Else
PathTemp=PathTemp & Arr_Path(Tempi) & "/"
End If
If CheckDir(PathTemp)=False Then
If MakeNewsDir(PathTemp)=False Then
SaveTf=False
Exit For
End If
End If
Next
End If

'ȥ���ظ�ͼƬ��ʼ
TempArray=Split(TempStr,"$Array$")
TempStr=""
For Tempi=0 To Ubound(TempArray)
If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
TempStr=TempStr & "$Array$" & TempArray(Tempi)
End If
Next
TempStr=Right(TempStr,Len(TempStr)-7)
TempArray=Split(TempStr,"$Array$")
'ȥ���ظ�ͼƬ����

'ת�����ͼƬ��ַ��ʼ
TempStr=""
For Tempi=0 To Ubound(TempArray)
TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
Next
TempStr=Right(TempStr,Len(TempStr)-7)
TempStr=Replace(TempStr,Chr(0),"")
TempArray2=Split(TempStr,"$Array$")
TempStr=""
'ת�����ͼƬ��ַ����

'ͼƬ�滻/����
Set Re = New Regexp
Re.IgnoreCase = True 
Re.Global = True

For Tempi=0 To Ubound(TempArray2)
RemoteFileUrl=TempArray2(Tempi)
If RemoteFileUrl<>"$False$" And SaveTf=True Then'����ͼƬ
ArrSaveFileName = Split(RemoteFileurl,".")
   strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'�ļ�����
If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
UploadFiles=""
ReplaceSaveRemoteFile=ConStr
Exit Function
End If

Randomize
RanNum=Int(900*Rnd)+100
   strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
Re.Pattern =TempArray(Tempi)
   If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
'********************************
PathTemp=SavePath & strFileName
ConStr=Re.Replace(ConStr,PathTemp)
Re.Pattern=strInstallDir & strChannelDir & "/"
UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")
Else
PathTemp=RemoteFileUrl
ConStr=Re.Replace(ConStr,PathTemp)
'UploadFiles=UploadFiles & "|" & RemoteFileUrl
End If
ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'������ͼƬ
Re.Pattern =TempArray(Tempi)
ConStr=Re.Replace(ConStr,RemoteFileUrl)
UploadFiles=UploadFiles & "|" & RemoteFileUrl
End If
Next 
Set Re=nothing
If UploadFiles<>"" Then
UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
End If
ReplaceSaveRemoteFile=ConStr
End function

'==================================================
'��������ReplaceSwfFile
'�� �ã���������·��
'�� ����ConStr ------ Ҫ�滻���ַ���
'�� ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSwfFile(ConStr,TistUrl)
If ConStr="$False$" or ConStr="" or TistUrl="" or TistUrl="$False$" Then
ReplaceSwfFile=ConStr
Exit Function
End If
Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

Set Re = New Regexp 
Re.IgnoreCase = True 
Re.Global = True
Re.Pattern ="<object.+?[^\>]>"
Set Matches =Re.Execute(ConStr) 
For Each Match in Matches
If TempStr<>"" then 
TempStr=TempStr & "$Array$" & Match.Value
Else
TempStr=Match.Value
End if
Next
If TempStr<>"" Then
TempArray=Split(TempStr,"$Array$")
TempStr=""
For Tempi=0 To Ubound(TempArray)
Re.Pattern ="value\s*=\s*.+?\.swf"
Set Matches =Re.Execute(TempArray(Tempi)) 
For Each Match in Matches
If TempStr<>"" then 
TempStr=TempStr & "$Array$" & Match.Value
Else
TempStr=Match.Value
End if
Next
Next
End if
If TempStr<>"" Then
Re.Pattern ="value\s*=\s*"
TempStr=Re.Replace(TempStr,"")
End If
If TempStr="" or IsNull(TempStr)=True Then
ReplaceSwfFile=ConStr
Exit function
End if
TempStr=Replace(TempStr,"""","")
TempStr=Replace(TempStr,"'","")
TempStr=Replace(TempStr," ","")

Set Matches=nothing
Set Re=nothing

'ȥ���ظ��ļ���ʼ
TempArray=Split(TempStr,"$Array$")
TempStr=""
For Tempi=0 To Ubound(TempArray)
If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
TempStr=TempStr & "$Array$" & TempArray(Tempi)
End If
Next
TempStr=Right(TempStr,Len(TempStr)-7)
TempArray=Split(TempStr,"$Array$")
'ȥ���ظ��ļ�����

'ת����Ե�ַ��ʼ
TempStr=""
For Tempi=0 To Ubound(TempArray)
TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
Next
TempStr=Right(TempStr,Len(TempStr)-7)
TempStr=Replace(TempStr,Chr(0),"")
TempArray2=Split(TempStr,"$Array$")
TempStr=""
'ת����Ե�ַ����

'�滻
Set Re = New Regexp
Re.IgnoreCase = True 
Re.Global = True
For Tempi=0 To Ubound(TempArray2)
RemoteFileUrl=TempArray2(Tempi)
Re.Pattern =TempArray(Tempi)
ConStr=Re.Replace(ConStr,RemoteFileUrl)
Next 
Set Re=nothing
ReplaceSwfFile=ConStr
End function

'==================================================
'��������SaveRemoteFile
'�� �ã�����Զ�̵��ļ�������
'�� ����LocalFileName ------ �����ļ���
'�� ����RemoteFileUrl ------ Զ���ļ�URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
SaveRemoteFile=True
  dim Ads,Retrieval,GetRemoteData
  Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
  With Retrieval
    .Open "Get", RemoteFileUrl, False, "", ""
    .Send
If .Readystate<>4 then
SaveRemoteFile=False
Exit Function
End If
    GetRemoteData = .ResponseBody
  End With
  Set Retrieval = Nothing
  Set Ads = Server.CreateObject("Adodb.Stream")
  With Ads
    .Type = 1
    .Open
    .Write GetRemoteData
    .SaveToFile server.MapPath(LocalFileName),2
    .Cancel()
    .Close()
  End With
  Set Ads=nothing
end Function

%>