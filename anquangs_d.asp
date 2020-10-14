<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link title="css" href="../css2012/index.css" rel="stylesheet"  type="text/css"/>
<LINK href="../css2012/menu.css" type=text/css rel=stylesheet>
<SCRIPT src="../css2012/menu.js" type=text/javascript></SCRIPT>
<SCRIPT language=javascript src="js/hhh.js"></SCRIPT>
<script language=javascript>
   <!--
    function CheckForm() {
      if(document.Login.UserName.value == '') {
        alert('请输入用户名！');
        document.Login.UserName.focus();
        return false;
      }
      if(document.Login.password.value == '') {
        alert('请输入密码！');
        document.Login.password.focus();
        return false;
      }
	  }
  //-->
</script>
</head>

<body>
<!--#include file="top.asp"-->

<DIV class="box960">
  <div class="boxl">
    <div class="t1">
      <div class="dq">当前位置：<a href="/">首页</a> > 安全短板公示&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>
	 <%  
dwt.out"<form method='Get' name='SearchForm' action='anquangs_luoshi.asp' >" & vbCrLf
	

'dwt.out "  <input type='Submit' name='Submit'  value='搜索'>" & vbCrLf

dwt.out "</form></div>" & vbCrLf

%> 
   
   
    <p class="br"> </p>
    <div class="boxlc boxlt">
      <%
	  
response.write "<table  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
response.write "<tr class=""title"">" 
response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center""><strong>序号</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""150""><div align=""center""><strong>短板内容</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""150""><div align=""center""><strong>整改措施</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>责任单位</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px"" width=""60px""><div align=""center""><strong>责任人</strong></div></td>"
response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>要求时间</strong><br><strong>完工时间<br>发布时间</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px""  width=""40px""><div align=""center""><strong>完成情况</strong></div></td>"
response.write "      <td width=""11%"" style=""border-bottom-style: solid;border-width:1px""><div align=""center""><strong>效果评价</strong></div></td>"
response.write "      <td style=""border-bottom-style: solid;border-width:1px""  width=""40px""><div align=""center""><strong>评价人</strong></div></td>"

response.write "    </tr>"

	if wangong=0 then sqlpxst="SELECT * from anquangs ORDER BY id DESC"
	if wangong=1 then sqlpxst="SELECT * from anquangs where isno=true ORDER BY id DESC"
	if wangong=2 then sqlpxst="SELECT * from anquangs where isno=false ORDER BY id DESC"



set rspxst=server.createobject("adodb.recordset")
rspxst.open sqlpxst,connaq,1,1
if rspxst.eof and rspxst.bof then 
response.write "<p align='center'>未添加</p>" 
else
           record=rspxst.recordcount
           if Trim(Request("PgSz"))="" then
               PgSz=20
           ELSE 
               PgSz=Trim(Request("PgSz"))
           end if 
           rspxst.PageSize = Cint(PgSz) 
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
           rspxst.absolutePage = page
           start=PgSz*Page-PgSz+1
           rowCount = rspxst.PageSize
           do while not rspxst.eof and rowcount>0
		   
			dim xh,xh_id
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			if xh_id mod 2 =1 then 
			  dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
                 dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&xh_id&"</div></td>"

'                 response.write "<tr class=""tdbg"" onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"
'                 response.write "     <td  style=""border-bottom-style: solid;border-width:1px"" ><div align=""center"">"&rspxst("id")&"</div></td>"
				 if rspxst("isno")=false then 
                 response.write "<td style=""border-bottom-style: solid;border-width:1px""  ><a href=anquangs_view.asp?id="&rspxst("id")&" target=_blank style=color:#FF0000>"&rspxst("pxst_title")&"</a></td>"
				 else
                 response.write "<td style=""border-bottom-style: solid;border-width:1px"" ><a href=anquangs_view.asp?id="&rspxst("id")&" target=_blank>"&rspxst("pxst_title")&"</a></td>"
				 
				 end if


wgsj="空"
if rspxst("wangong_date")<>"" then wgsj=rspxst("wangong_date")
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_zgcs")&"&nbsp;</div></td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("zr_danwei")&"&nbsp;</div></td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("zr_ren")&"&nbsp;</div></td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("yaoqiu_date")&"&nbsp;</div><div align=""center"">"&wgsj&"&nbsp;</div><div align=""center"">"&rspxst("pxst_date")&"</div></td>"
                 'response.write "<td style=""border-bottom-style: solid;border-width:1px;color:#FF0000""><div align=""center"">"
				if rspxst("isno")=true then
				 response.write "<td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">已完工</div></td>"
				else
				 response.write "<td  style=""border-bottom-style: solid;border-width:1px;color:#FF0000""><div align=""center"">未完工</div></td>"
				
				end if
				
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_estimation")&"&nbsp;</div></td>"
                 response.write "      <td  style=""border-bottom-style: solid;border-width:1px""><div align=""center"">"&rspxst("pxst_estimator")&"&nbsp;</div></td>"
				'response.write  "</div></td>"
                 response.write "    </tr>"
                 RowCount=RowCount-1
          rspxst.movenext
          loop
       end if
       rspxst.close
       set rspxst=nothing
        conn.close
        set conn=nothing
        response.write "</table>"
     url="anquangs_d.asp?action="
 
				
if request("news_class")="" then        
   call showpage1(page,url,total,record,PgSz)
else 
call showpage(page,url,total,record,PgSz)
	end if   
						
				
				
				
				

		%>
    </div>
  </div>
  
    <!--#include file="index_left.asp"--> 

</div>
<div class="clear"></div>
<div class=miniNav>
  <div class="box960" align="center"><br>
    <br>
    <b>设备管理系统</b> <br>
    <br>
  </div>
</div>
</body>
</html>
