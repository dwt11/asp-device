
<DIV class="boxr">
  <div class="boxs">
    <DIV class=hd><span>�û���¼</span></DIV>
    <DIV class=bd>
      <%
	   if session("UserName")<>"" then 
	    %>
      <div style="padding-left:10px;padding-right:10px"><span id="webasp_time"></span><script>setInterval("webasp_time.innerHTML=new Date().toLocaleString()+' ����'+'��һ����������'.charAt (new Date().getDay());",1000);</script>
        <%
		 
		  dwt.out "<br><b>"&session("UserName1")&"</b>,����</br>"
		  dwt.out "<a href=/main.asp>������Ϣ����̨����</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='/login.asp?action=Logout'>�˳�</a>"
       dwt.out "</div>"
		  	   else
	   
	   
	   %>
        <form name='Login' action='../login.asp' method='post' target='_parent'  onSubmit='return CheckForm();'>
          <div class="fb"><span>�û���:</span>
            <input name='UserName' type='text' id='UserName' class="ipt-txt" size="20" value="admin">
          </div>
          <div class="fb"><span>����:</span>
            <input name='password' type='password' id='Password' class="ipt-txt" size="21" value="123456">
          </div>
          <div class="submit">
            <button type="submit" class="btn-1">��¼</button>
            <input type='hidden' name='Action' value='Login'>
          </div>
        </form>
        <%end if %>
      </DIV>
    </DIV>
  </DIV>
  <div class="boxs">
    <DIV class=hd><span>֪ͨ</span></DIV>
    <DIV class=bd>
      <DIV class=innerBox>
        <UL class=toplist>
          <%i =0
sqlqxtb="SELECT top 5 * from xzgl_news where news_class=52  ORDER BY id DESC"
'sqlqxtb="SELECT top 6 * from xzgl_news where news_class=52 AND MONTH(NEWS_DATE)=MONTH(NOW()) AND YEAR(NEWS_DATE)=YEAR(NOW()) ORDER BY id DESC"
set rsqxtb=server.createobject("adodb.recordset")
rsqxtb.open sqlqxtb,conna,1,1
if rsqxtb.eof and rsqxtb.bof then 
'Dwt.out "<p align='center'>δ�������</p>" 
else
dwt.out "<marquee  scrollamount=2 height=100 onmouseover=stop()  onmouseout=start() direction='up'>"
do while not rsqxtb.eof
title=rsqxtb("news_title")
if len(title)>35 then
title=left(title,25)&"..."

%>
          <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" title="<%=rsqxtb("news_title")%>" target=_blank><%=title%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rsqxtb("news_date")%>]
            <%else%>
          
          <li><a href="news_view.asp?ID=<%=rsqxtb("id")%>" target=_blank><%=rsqxtb("news_title")%></a>&nbsp;&nbsp;&nbsp;&nbsp;[<%=rsqxtb("news_date")%>]
            <%end if
					i=i+1
'if i=8 then exit do
rsqxtb.movenext
loop
end if
					dwt.out "</marquee>"
rsqxtb.close
set rsqxtb=nothing
%>
        </UL>
      </DIV>
    </DIV>
  </DIV>
  
  
</DIV>
