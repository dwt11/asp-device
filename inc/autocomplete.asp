<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<script language="javascript" runat="server">
function decode(str) {
        return unescape(str);
}
</script>
<%

typing =decode(trim(Request.QueryString("typing")))
'conntext=trim(Request.QueryString("conntext")) '�����ַ���
dbname=trim(Request.QueryString("dbname"))   ';���ݿ����� 
zdtext=trim(Request.QueryString("zdtext"))   'Ҫ��ȡ���ֶ�
btext=trim(Request.QueryString("btext"))    '������
Response.ContentType = "text/html"
Response.Charset = "GB2312"   '�����������


url="/inc/autocomplete.asp?dbname="&dbname&"&zdtext="&zdtext&"&btext="&btext&"&typing="&typing
db_path = "/ybdata/"&dbname&".mdb"
Set conn= Server.CreateObject("ADODB.Connection")
connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(db_path)
conn.Open connstr


sql="SELECT  distinct "&zdtext&" FROM "&btext&" WHERE "&zdtext&" LIKE '"&typing&"%'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then 
else



	record=rs.recordcount
   if Trim(Request("PgSz"))="" then
	   PgSz=10
   ELSE 
	   PgSz=Trim(Request("PgSz"))
   end if 
   rs.PageSize = Cint(PgSz) 
   total=int(record/PgSz*-1)*-1
   page=Request("page")
   if page="" Then
	  page = 1
   else
	 page=page+1
	 page=page-1
   end if
   if page<1 Then 
	  page=1
   end if
   rs.absolutePage = page
   start=PgSz*Page-PgSz+1
   rowCount = rs.PageSize

do while  NOT rs.EOF and rowcount>0
 
			dim xh_id
				xh_id=((page-1)*pgsz)+1+xh
				xh=xh+1
 response.write "<div onselect='autoback("","&replace(rs(0),"""","")&");' onfocus='update(this,"""&replace(rs(0),"""","")&""")'>"
 'response.write "	<span class='informal'>["&rs(2)&"]</span>"   '������ʾ����ʾ,�ݲ���
 response.write "<span class='green'>"&xh_id&"</span> "&replace(rs(0),"""","")
 response.write "</div>"
 RowCount=RowCount-1
 rs.MoveNext
loop
       call showpage(page,url,total,record,PgSz)
 conn.close

 set rs=nothing
 set conn=nothing
end if 



'********************************************8
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
'URL�д�����
'*******************************************
sub showpage(page,url,total,record,pgsz)
   response.write "<div align='center'>"
   response.write"<span style='color:red'>"&page&"</span>/"&total&"&nbsp;&nbsp;"
   'response.write record&"��<br/>"
   if page="" then page=1
   if page > 1 Then 
      'response.write "<a href="&url&"&page=1>��ǰ&nbsp;"
	  response.write"<a href="&url&"&pgsz="&pgsz&"&page="&page-1&">��һҳ</a>&nbsp;"
   else
      response.write ""
   end if 
   if RowCount = 0 and page <>Total then 
     response.write "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&">��һҳ</a>"
	 '" <a href="&url&"&pgsz="&pgsz&"&page="&total&">���</a>"
   else
     response.write ""
   end if
   response.write "</div>"
end sub




 %>           
     