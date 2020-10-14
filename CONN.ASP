<%
dim conn,conna,connb,connc,connd,conne,conndcs,connkc,connjg,connaq,connj
dim db,xzgldb,scgldb,bugdb,messagedb,pxjhzjdb,dcsghjxdb,kcgldb,jgtzdb,dgtzldb,aqgldb
dim connstr,connxzgl,connscgl,connbug,connmessage,connpxjhzj,conndcsghjx,connjgtz,conndgtzl,connaqgl,connjlgl
dim connkcgl,jlgl
dim connzj,connzjtz,zjtzdb
dim connzb,connzblog,zblogdb,csyydb,conncsyy,conncs,conndgt
   'on error resume next


'data.mdb 数据库包括技术档案、用户管理
db="/ybdata/data.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
    conn.Open connstr



dim dbwdjc,connw,connwdjc
dbwdjc="/ybdata/wdjc.mdb"
Set connw = Server.CreateObject("ADODB.Connection")
Connwdjc = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(dbwdjc)
    connw.Open connwdjc

'xzgldb.mdb 数据库包括月计划、月总结\新闻
xzgldb="/ybdata/xzgldb.mdb"
Set conna = Server.CreateObject("ADODB.Connection")
connxzgl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(xzgldb)
    conna.Open connxzgl
	
'aqgldb.mdb 数据库包括安全月计划、月总结
aqgldb="/ybdata/aqgldb.mdb"
Set connaq= Server.CreateObject("ADODB.Connection")
connaqgl= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(aqgldb)
    connaq.Open connaqgl



'scgldb.mdb 数据库生产管理
scgldb="/ybdata/scgldb.mdb"
Set connb = Server.CreateObject("ADODB.Connection")
connscgl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(scgldb)
    connb.Open connscgl

'bug.mdb BUG收集
bugdb="/ybdata/bug.mdb"
Set connc = Server.CreateObject("ADODB.Connection")
connbug = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(bugdb)
    connc.Open connbug

'message.mdb 消息传送
messagedb="/ybdata/message.mdb"
Set connd = Server.CreateObject("ADODB.Connection")
connmessage = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(messagedb)
    connd.Open connmessage

'pxjhzj.mdb 数据库包括技术档案、用户管理
pxjhzjdb="/ybdata/pxjhzj.mdb"
Set conne = Server.CreateObject("ADODB.Connection")
Connpxjhzj = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(pxjhzjdb)
    conne.Open connpxjhzj

'dcsghjx.mdb DCS检修/更换计划，检修记录
dcsghjxdb="/ybdata/dcsghjx.mdb"
   
Set conndcs = Server.CreateObject("ADODB.Connection")
conndcsghjx = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(dcsghjxdb)
    conndcs.Open conndcsghjx

'dcsghjx.mdb DCS检修/更换计划，检修记录
kcgldb="/ybdata/kcgl.mdb"
Set connkc = Server.CreateObject("ADODB.Connection")
connkcgl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(kcgldb)
    connkc.Open connkcgl

jgtzdb="/ybdata/jgtz.mdb"
Set connjg = Server.CreateObject("ADODB.Connection")
connjgtz = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(jgtzdb)
    connjg.Open connjgtz

zjtzdb="/ybdata/zjtz.mdb"
Set connzj = Server.CreateObject("ADODB.Connection")
connzjtz= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(zjtzdb)
    connzj.Open connzjtz

	zblogdb="/ybdata/zblog.mdb"
	Set connzb = Server.CreateObject("ADODB.Connection")
	connzblog= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(zblogdb)
	connzb.Open connzblog

'sub conn_ndjx()
    dim ndjxdb,connndjx,connnd
	ndjxdb="/ybdata/ndjx1.mdb"
	Set connnd = Server.CreateObject("ADODB.Connection")
	connndjx= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ndjxdb)
	connnd.Open connndjx
'end sub

        dim leftmdb,connleft,connl
		leftmdb="/ybdata/left.mdb"
		Set connleft = Server.CreateObject("ADODB.Connection")
		connl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(leftmdb)
		connleft.Open connl  
		

csyydb="/ybdata/csyy.mdb"
Set conncs = Server.CreateObject("ADODB.Connection")
conncsyy = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(csyydb)
    conncs.Open conncsyy

dgtzldb="/ybdata/dgtzl.mdb"
Set conndgt = Server.CreateObject("ADODB.Connection")
conndgtzl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(dgtzldb)
    conndgt.Open conndgtzl

'jlgl.mdb 数据库为计量管理的内容
jlgl="/ybdata/jlgl.mdb"
Set connj = Server.CreateObject("ADODB.Connection")
connjlgl = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(jlgl)
    connj.Open connjlgl
		  
Sub CloseConn()
    On Error Resume Next
    If IsObject(conn) Then
        conn.Close
        Set conn = Nothing
    End If
    If IsObject(conna) Then
        conna.Close
        Set conna = Nothing
    End If
    If IsObject(connb) Then
        connb.Close
        Set connb = Nothing
    End If
    If IsObject(connc) Then
        connc.Close
        Set connc = Nothing
    End If
    If IsObject(connd) Then
        connd.Close
        Set connd = Nothing
    End If

 If IsObject(conne) Then
        conne.Close
        Set conne = Nothing
    End If

 If IsObject(conndcs) Then
        conndcs.Close
        Set conndcs = Nothing
    End If
    
     If IsObject(connkc) Then
        connkc.Close
        Set connkc = Nothing
    End If

     If IsObject(connjg) Then
        connjg.Close
        Set connjg = Nothing
    End If

If IsObject(connzb) Then
        connzb.Close
        Set connzb = Nothing
    End If

If IsObject(connj) Then
        connj.Close
        Set connj = Nothing
    End If

End Sub

%>