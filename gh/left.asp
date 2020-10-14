  <DIV class="boxr">
    <div class="boxs">
      <DIV class=hd><span>用户登录</span></DIV>
      <DIV class=bd>
       <%
	   if session("UserName")<>"" then 
	    %><div style="padding-left:10px;padding-right:10px"><span id="webasp_time"></span><script>setInterval("webasp_time.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt (new Date().getDay());",1000);</script>
		<%
		 
		  dwt.out "<br><b>"&session("UserName1")&"</b>,您好</br>"
		  dwt.out "<a href=/main.asp>进入信息网后台管理</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='/login.asp?action=Logout'>退出</a>"
       dwt.out "</div>"
		  	   else
	   
	   
	   %>
       
       
        <form name='Login' action='../login.asp' method='post' target='_parent'  onSubmit='return CheckForm();'>
          <div class="fb"><span>用户名:</span>
            <input name='UserName' type='text' id='UserName' class="ipt-txt" size="20">
          </div>
          <div class="fb"><span>密码:</span>
            <input name='password' type='password' id='Password' class="ipt-txt" size="21">
          </div>
          <div class="submit">
            <button type="submit" class="btn-1">登录</button>
            <input type='hidden' name='Action' value='Login'>
          </div>
        </form>
      
      <%end if %>
      
      </DIV>
      
      
    </DIV>
    <div class="boxs">
      <DIV class=hd><span>栏目分类</span></DIV>
      <DIV class=bd>
        <DIV class=innerBox>
          <UL class=toplist>
           <%dim i
		sqltree="SELECT * from dgtzl_index_gh where index=0 order by orderby"& vbCrLf
		set rstree=server.createobject("adodb.recordset")
		rstree.open sqltree,conndgt,1,1
		do while not rstree.eof
			dim urltmp
			
			sqltree2="SELECT * from dgtzl_index_gh where index="&rstree("id")&" order by orderby"& vbCrLf
			set rstree2=server.createobject("adodb.recordset")
			rstree2.open sqltree2,conndgt,1,1
			if not rstree2.eof then 
			    urltmp="index.asp?classid="&rstree("id")
			else
				urltmp="showlist.asp?classid="&rstree("id")
			end if 
			
			
			
			
			dwt.out "<li><a href="&urltmp&">"&rstree("class_name")&"</a></li>" & vbCrLf
		rstree.movenext
		loop
		rstree.close
		set rstree=nothing
		%>
          </UL>
        </DIV>
      </DIV>
    </DIV>
  </DIV>
