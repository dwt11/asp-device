<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'dim starttime : starttime=timer
%>
<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"-->

<%
keys=trim(request("keyword")) 
sscjid=trim(request("sscj")) 
ssghid=trim(request("ssgh")) 
sb_classid = Trim(Request("sbclassid"))
   if sb_classid="" then sb_classid=24
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sb_classid)(0)

Dwt.out "<html>"& vbCrLf
Dwt.out "<head>" & vbCrLf
Dwt.out "<title>������������ҳ</title>"& vbCrLf
Dwt.out "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/grid.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/docs.css' rel='stylesheet' type='text/css'>"& vbCrLf

Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out"<script language=javascript src='/js/popselectdate.js'></script>"

Dwt.out "</head>"& vbCrLf
Dwt.out "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>"& vbCrLf

%>
<!--<Div id="loading">
  <Div class="loading-indicator" ><img src="img_ext/default/grid/loading.gif" style="width:16px;height:16px;" align="absmiddle"> ҳ�������...</Div>
</Div>
--><%
action=request("action")
select case action
  case "add"
      if truepagelevelh(session("groupid"),1,session("pagelevelid")) then call add'����豸����ѡ��
  case "img"
    dwt.out "<br/><Div align=center><b>"&conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&request("sbid"))(0)&" ͼƬ��Ϣ</b></div><br/>"
	dwt.out conn.Execute("SELECT sb_img FROM sb WHERE sb_id="&request("sbid"))(0)
  case "addsb"
      call addsb'ѡ����������豸ҳ��
  case "saveaddsb"
      call saveaddsb'�豸��ӱ���
  case "edit"
      if truepagelevelh(session("groupid"),2,session("pagelevelid")) then call edit
  case "saveedit"'�༭�ӷ���
      call saveedit'�༭�����ӷ���
  case "del"
        if truepagelevelh(session("groupid"),3,session("pagelevelid")) then 
			
	    Dwt.savesl "�豸����-"&zclass(conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&request("id"))(0)),"ɾ��",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&request("id"))(0)

			Set Rs = Server.CreateObject("ADODB.Recordset")
			Sql = "Delete From sb Where sb_id="&request("id")
			Conn.execute(Sql)
			Dwt.out "<Script Language=Javascript>history.back()</Script>"
			set rs=nothing
			set conn=nothing
		end if 
  case ""
      if truepagelevelh(session("groupid"),0,session("pagelevelid")) then call main
end select	  	 

sub add()
	 Dwt.out "<SCRIPT language=javascript>" & vbCrLf
	Dwt.out "function checkadd(){" & vbCrLf
	
	Dwt.out " if(document.form1.sb_class.value==0){" & vbCrLf
	Dwt.out "      alert('�豸һ������δѡ��');" & vbCrLf
	Dwt.out "   document.form1.sb_class.focus();" & vbCrLf
	Dwt.out "      return false;" & vbCrLf
	Dwt.out "    }" & vbCrLf
	Dwt.out " if(document.form1.sbclassid.value==0){" & vbCrLf
	Dwt.out "      alert('�豸��������δѡ��');" & vbCrLf
	Dwt.out "   document.form1.sbclassid.focus();" & vbCrLf
	Dwt.out "      return false;" & vbCrLf
	Dwt.out "    }" & vbCrLf
	
	
	
	Dwt.out "    }" & vbCrLf
	Dwt.out "</SCRIPT>" & vbCrLf
		Dwt.out"<Div align=center><Div style='WIDTH: 480px;padding-top:100px'>"& vbCrLf
		Dwt.out"  <Div class=x-box-tl>"& vbCrLf
		Dwt.out"	<Div class=x-box-tr>"& vbCrLf
		Dwt.out"	  <Div class=x-box-tc></Div>"& vbCrLf
		Dwt.out"	</Div>"& vbCrLf
		Dwt.out"  </Div>"& vbCrLf
		Dwt.out"  <Div class=x-box-ml>"& vbCrLf
		Dwt.out"	<Div class=x-box-mr>"& vbCrLf
		Dwt.out"	  <Div class=x-box-mc>"& vbCrLf
		Dwt.out"		<H3 style='MARGIN-BOTTOM: 5px'>����豸</H3>"& vbCrLf
		Dwt.out"		<Div id=form-ct>"& vbCrLf
		Dwt.out "<form method='post' class='x-form' action='sb.asp' name='form1' onsubmit='javascript:return checkadd();' >"
		Dwt.out"			<Div class='x-form-ct'>"& vbCrLf
		Dwt.out"							<Div class='x-form-item'>"& vbCrLf
		Dwt.out"				<LABEL style='WIDTH: 105px'>ѡ���豸����:</LABEL>"& vbCrLf
		Dwt.out"				<Div class='x-form-element' >"& vbCrLf
		
		
		
		Dwt.out "<select name='sb_class' size='1' id='cat1' onChange=""selectpc(this.value,'b',document.form1.sbclassid)"">"
		Dwt.out "  <option selected value='0'>ѡ��һ������</option>"
		sql="SELECT * from sbclass where sbclass_zclass=0 "& vbCrLf
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		do while not rs.eof
			Dwt.out"<option value='"&rs("sbclass_id")&"'>"&rs("sbclass_name")&"</option>"& vbCrLf
			rs.movenext
		loop
		rs.close
		set rs=nothing
		Dwt.out "</select>"
		Dwt.out "<select name='sbclassid' size='1' id='cat2' >"
		Dwt.out "  <option selected value=0>ѡ���������</option>"
		Dwt.out "</select></td></tr>"& vbCrLf
		Dwt.out "<script language='javascript'>"& vbCrLf
		Dwt.out "function selectpc(parentValue,child,addObj){"& vbCrLf
	
	
	dim b,bv,b_p
		sql="SELECT * from sbclass where sbclass_zclass=0 "& vbCrLf
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
			 b="var b =   new Array("
			bv="var bv =   new Array("
			b_p="var b_p =   new Array("
	   
		do while not rs.eof
			sqlz="SELECT * from sbclass where sbclass_zclass="&rs("sbclass_id")
			set rsz=server.createobject("adodb.recordset")
			rsz.open sqlz,conn,1,1
			if rsz.eof and rsz.bof then
			   b=b&"'�޶�������',"
			   bv=bv&"'0',"
			   b_p=b_p&"'"&rs("sbclass_id")&"',"
			else
			do while not rsz.eof
				b=b&"'"&rsz("sbclass_name")&"',"
				bv=bv&"'"&rsz("sbclass_id")&"',"
				b_p=b_p&"'"&rs("sbclass_id")&"',"
	
	
			   rsz.movenext
			loop
			end if 
			rsz.close
			set rsz=nothing
	
			rs.movenext
		loop
		rs.close
		set rs=nothing
		b=left(b,len(b)-1)
		bv=left(bv,len(bv)-1)
		b_p=left(b_p,len(b_p)-1)
		b=b&");"
		bv=bv&");"
		b_p=b_p&");"
		Dwt.out b & vbCrLf
		Dwt.out bv & vbCrLf
		Dwt.out b_p & vbCrLf
		
		
		
		Dwt.out "var labelValue = new Array();"& vbCrLf
		Dwt.out "var labelText =  new Array();"& vbCrLf
		Dwt.out "var k = 0;"& vbCrLf
		
		Dwt.out "cObj = eval(child);"& vbCrLf
		Dwt.out "cObjV = eval(child+'v');"& vbCrLf
		Dwt.out "cpObj = eval(child + '_p');"& vbCrLf
		Dwt.out "for(i=0; i<cpObj.length; i++)"& vbCrLf
		Dwt.out "{"& vbCrLf
		Dwt.out "	if(cpObj[i] == parentValue)"& vbCrLf
		Dwt.out "	{"& vbCrLf
		Dwt.out "		labelText[k] =  cObj[i];"& vbCrLf
		Dwt.out "		labelValue[k] =	cObjV[i]; "& vbCrLf
		Dwt.out "		k++;"& vbCrLf
		Dwt.out "	}"& vbCrLf
		Dwt.out "}"& vbCrLf
		
		
		Dwt.out "addObj.options.length = 0;"& vbCrLf
		Dwt.out "addObj.options[0] = new Option('==ѡ���������==','0');"& vbCrLf
		Dwt.out "for(i = 0; i < labelText.length; i++) {"& vbCrLf
		Dwt.out "	addObj.add(document.createElement('option'));"& vbCrLf
		Dwt.out "	addObj.options[i+1].text=labelText[i];"& vbCrLf
		Dwt.out "	addObj.options[i+1].value=labelValue[i];"& vbCrLf
		Dwt.out "}"& vbCrLf
		Dwt.out "addObj.selectedIndex = 0;"& vbCrLf
	Dwt.out "}"& vbCrLf
	Dwt.out "</script>"& vbCrLf
		
		
		
		Dwt.out"				</Div>"& vbCrLf
		Dwt.out"			  </Div>"& vbCrLf
		Dwt.out"			  <Div class=x-form-clear-left></Div>"& vbCrLf
		
		Dwt.out"			  <Div class=x-form-clear></Div>"& vbCrLf
		Dwt.out"			</Div>"& vbCrLf
		Dwt.out"			<Div class=x-form-btns-ct>"& vbCrLf
		Dwt.out"			  <Div class='x-form-btns x-form-btns-center'>"& vbCrLf
		Dwt.out"			  <input name='action' type='hidden' value='addsb'>    <input  type='submit' name='Submit' value=' ��һ�� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'>"& vbCrLf
		Dwt.out"				<Div class=x-clear></Div>"& vbCrLf
		Dwt.out"			  </Div>"& vbCrLf
		Dwt.out"			</Div>"& vbCrLf
		Dwt.out"		  </FORM>"& vbCrLf
		Dwt.out"		</Div>"& vbCrLf
		Dwt.out"	  </Div>"& vbCrLf
		Dwt.out"	</Div>"& vbCrLf
		Dwt.out"  </Div>"& vbCrLf
		Dwt.out"  <Div class=x-box-bl>"& vbCrLf
		Dwt.out"	<Div class=x-box-br>"& vbCrLf
		Dwt.out"	  <Div class=x-box-bc></Div>"& vbCrLf
		Dwt.out"	</Div>"& vbCrLf
		Dwt.out"  </Div>"& vbCrLf
		Dwt.out"</Div>"& vbCrLf
		Dwt.out"</Div> "& vbCrLf  
		
	   
	   
end sub


sub addsb()
'sbclass_id=request("sbclassid")
	%>
<!--	<Div style="WIDTH: 300px">
  <Div class=x-box-tl>
    <Div class=x-box-tr>
      <Div class=x-box-tc></Div>
    </Div>
  </Div>
  <Div class=x-box-ml>
    <Div class=x-box-mr>
      <Div class=x-box-mc>
        <%'Dwt.out "<H3 style='MARGIN-BOTTOM: 5px'><strong>�����豸: "&sb_classname&"</strong></H3>"%>
        <Div id=form-ct>
          <FORM class=" x-form" id=ext-gen25 method=post>
            <Div class=x-form-ct id=ext-gen24>
              <Div class="x-form-item ">
                <LABEL style="WIDTH: 75px" for=ext-comp-1001>First Name:</LABEL>
                <Div class=x-form-element id=x-form-el-ext-comp-1001 style="PADDING-LEFT: 80px">
                  <INPUT class=" x-form-text x-form-field" id=ext-comp-1001 style="WIDTH: 175px" name=first autocomplete="off">
                </Div>
              </Div>
              <Div class=x-form-clear-left></Div>
              <Div class="x-form-item ">
                <LABEL style="WIDTH: 75px" for=ext-comp-1002>Last Name:</LABEL>
                <Div class=x-form-element id=x-form-el-ext-comp-1002 style="PADDING-LEFT: 80px">
                  <INPUT class=" x-form-text x-form-field" id=ext-comp-1002 style="WIDTH: 175px" name=last autocomplete="off">
                </Div>
              </Div>
              <Div class=x-form-clear-left></Div>
              <Div class="x-form-item ">
                <LABEL style="WIDTH: 75px" for=ext-comp-1003>Company:</LABEL>
                <Div class=x-form-element id=x-form-el-ext-comp-1003 style="PADDING-LEFT: 80px">
                  <INPUT class=" x-form-text x-form-field" id=ext-comp-1003 style="WIDTH: 175px" name=company autocomplete="off">
                </Div>
              </Div>
              <Div class=x-form-clear-left></Div>
              <Div class="x-form-item ">
                <LABEL style="WIDTH: 75px" for=ext-comp-1004>Email:</LABEL>
                <Div class=x-form-element id=x-form-el-ext-comp-1004 style="PADDING-LEFT: 80px">
                  <INPUT class=" x-form-text x-form-field" id=ext-comp-1004 style="WIDTH: 175px" name=email autocomplete="off">
                </Div>
              </Div>
              <Div class=x-form-clear-left></Div>
              <Div class=x-form-clear id=ext-gen27></Div>
            </Div>
            <Div class=x-form-btns-ct>
              <Div class="x-form-btns x-form-btns-center">
                <TABLE cellSpacing=0>
                  <TBODY>
                    <TR>
                      <TD class=x-form-btn-td id=ext-gen45><TABLE class="x-btn-wrap x-btn " id=ext-gen46 style="WIDTH: 75px" cellSpacing=0 cellPadding=0 border=0>
                          <TBODY>
                            <TR>
                              <TD class=x-btn-left><I>&nbsp;</I></TD>
                              <TD class=x-btn-center><EM unselectable="on">
                                <BUTTON class=x-btn-text id=ext-gen47>Save</BUTTON>
                                </EM></TD>
                              <TD class=x-btn-right><I>&nbsp;</I></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                      <TD class=x-form-btn-td id=ext-gen54><TABLE class="x-btn-wrap x-btn" id=ext-gen55 style="WIDTH: 75px" cellSpacing=0 cellPadding=0 border=0>
                          <TBODY>
                            <TR>
                              <TD class=x-btn-left><I>&nbsp;</I></TD>
                              <TD class=x-btn-center><EM unselectable="on">
                                <BUTTON class=x-btn-text id=ext-gen56>Cancel</BUTTON>
                                </EM></TD>
                              <TD class=x-btn-right><I>&nbsp;</I></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <Div class=x-clear></Div>
              </Div>
            </Div>
          </FORM>
        </Div>
      </Div>
    </Div>
  </Div>
  <Div class=x-box-bl>
    <Div class=x-box-br>
      <Div class=x-box-bc></Div>
    </Div>
  </Div>
</Div>
--><%	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.sb_sscj.value==''){" & vbCrLf
Dwt.out "      alert('��ѡ���������䣡');" & vbCrLf
Dwt.out "   document.form.sb_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_ssgh.value==0){" & vbCrLf
Dwt.out "      alert('��ѡ������װ�ã�');" & vbCrLf
Dwt.out "   document.form.sb_ssgh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_wh.value==''){" & vbCrLf
Dwt.out "      alert('����дλ�ţ�');" & vbCrLf
Dwt.out "   document.form.sb_wh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf


Dwt.out " if(document.form.sb_sccj.value==''){" & vbCrLf
Dwt.out "      alert('����д�������ң�');" & vbCrLf
Dwt.out "   document.form.sb_sccj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out"<form method='post' action='sb.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	Dwt.out"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.out"<tr class='title'>"& vbCrLf
	Dwt.out"<td height='22' colspan='2'><Div align=center><strong>���� "&sb_classname&" �豸</strong></Div></tr>"& vbCrLf
	
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	if session("level")=0 then 
	'����˵��������levelname���ж�ȡȫ����levelclass=1�ĳ������ƣ�Ȼ����ݳ���ID��bzname���ж�ȡ��Ӧ�İ���������ʾ
	
	Dwt.out"<select name='sb_sscj' size='1'  onChange=""redirect(this.options.selectedIndex)"">"& vbCrLf
    Dwt.out"<option  selected>ѡ����������</option>"& vbCrLf
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    do while not rscj.eof
       	Dwt.out"<option value='"&rscj("levelid")&"'>"&rscj("levelname")&"</option>"& vbCrLf
	
		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out"</select>"  	 & vbCrLf
    Dwt.out "<select name='sb_ssgh' size='1' >" & vbCrLf
    Dwt.out "<option  selected>ѡ��װ�÷���</option>" & vbCrLf
    Dwt.out "</select></td></tr>  "  & vbCrLf
    Dwt.out "<script><!--" & vbCrLf
    Dwt.out "var groups=document.form.sb_sscj.options.length" & vbCrLf
    Dwt.out "var group=new Array(groups)" & vbCrLf
    Dwt.out "for (i=0; i<groups; i++)" & vbCrLf
    Dwt.out "group[i]=new Array()" & vbCrLf
    Dwt.out "group[0][0]=new Option(""ѡ��װ�÷���"","" "");" & vbCrLf
	
	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
    set rscj=server.createobject("adodb.recordset")
    rscj.open sqlcj,conn,1,1
    
	do while not rscj.eof
     ii=0		
		sqlbz="SELECT * from ghname where sscj="&rscj("levelid")
        set rsbz=server.createobject("adodb.recordset")
        rsbz.open sqlbz,conn,1,1
        if rsbz.eof and rsbz.bof then
		   Dwt.out "group["&rscj("levelid")&"][0]=new Option(""��װ�÷���"",""0"");" & vbCrLf
		else
		do while not rsbz.eof
		   'Dwt.out"group["&rsbz("sscj")&"][0]=new Option(""����"",""0"");" & vbCrLf
		   Dwt.out"group["&rsbz("sscj")&"]["&ii&"]=new Option("""&rsbz("gh_name")&""","""&rsbz("ghid")&""");" & vbCrLf
		  ii=ii+1
		   rsbz.movenext
	    loop
	    end if 
		rsbz.close
	    set rsbz=nothing

		rscj.movenext
	loop
	rscj.close
	set rscj=nothing
    Dwt.out "var temp=document.form.sb_ssgh" & vbCrLf
    Dwt.out "function redirect(x){" & vbCrLf
    Dwt.out "for (m=temp.options.length-1;m>0;m--)" & vbCrLf
    Dwt.out "temp.options[m]=null" & vbCrLf
    Dwt.out "for (i=0;i<group[x].length;i++){" & vbCrLf
    Dwt.out "temp.options[i]=new Option(group[x][i].text,group[x][i].value)" & vbCrLf
    Dwt.out "}" & vbCrLf
    Dwt.out "temp.options[0].selected=true" & vbCrLf
    Dwt.out "}//--></script>" & vbCrLf



  else 	 
   Dwt.out"<input name='sb_sscj' type='text' value='"&sscjh(session("levelclass"))&"'  disabled='disabled' >"& vbCrLf
   Dwt.out"<input name='sb_sscj' type='hidden' value="&session("levelclass")&">"& vbCrLf
   if session("levelclass")=4 then 
      sql="SELECT * from ghname "
   else
      sql="SELECT * from ghname where sscj="&session("levelclass")
   end if 
   set rs=server.createobject("adodb.recordset")
   rs.open sql,conn,1,1
   Dwt.out"<select name='sb_ssgh' size='1'>"
   
   if rs.eof and rs.bof then 
   	  Dwt.out"<option value='0'>δ���װ��</option>"
   else   
	  'Dwt.out"<option value='0'>����</option>"
      do while not rs.eof
	     Dwt.out"<option value='"&rs("ghid")&"'>"&rs("gh_name")&"</option>"
	  rs.movenext
      loop
	  end if 
	 Dwt.out"</select>" 
  rs.close
  set rs=nothing
 end if 
    Dwt.out"</td></tr>  "  	 

	
	
	if zclassor(sb_classid) then
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>���ͣ� </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass sb_classid,0
		Dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>λ�ţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_wh' type='text' ></td></tr>"& vbCrLf
	
'	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�豸���ԣ� </strong></td>"   & vbCrLf   
'	Dwt.out"<td width='60%' class='tdbg'>"
'	Dwt.out" <label><input type='checkbox' name='sb_isls'/>�Ƿ����� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_iszj'/>�Ƿ��ܼ� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_isbw'/>�Ƿ��� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_isjl'/>�Ƿ�������� </label>"
	
	Dwt.out "</td></tr>"& vbCrLf
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��ã� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_whqk' size='1' >"
	Dwt.out"<option value='0'"
	
	Dwt.out">��ѡ��������</option>"
	Dwt.out"<option value='1'>���</option>"
	Dwt.out"<option value='2'>�����</option>"
	Dwt.out"</select></td></tr>"
	if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&sb_classid)(0) then 
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>׼ȷ�� </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_zqqk' size='1' >"
		Dwt.out"<option value='0'>��ѡ��׼ȷ���</option>"
		Dwt.out"<option value='1'>�����С</option>"
		Dwt.out"<option value='2'>�м�</option>"
		Dwt.out"<option value='3'>>95%</option>"
		Dwt.out"</select></td></tr>"
   end if 
	
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>Ͷ�ˣ� </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_tyqk' size='1' >"
		Dwt.out"<option value='0'>��ѡ��Ͷ�����</option>"
		Dwt.out"<option value='1'>Ͷ��</option>"
		Dwt.out"<option value='2'>ԭ��δͶ��</option>"
		Dwt.out"<option value='3'>����ԭ��δͶ��</option>"
		Dwt.out"</select></td></tr>"
   
   
   	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�ּ��� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_fj' size='1' >"
	Dwt.out"<option value='0'>��ѡ��ּ�</option>"
	Dwt.out"<option value='1'>��</option>"
	Dwt.out"<option value='2'>���</option>"
	Dwt.out"<option value='3'>����</option>"
	Dwt.out"</select></td></tr>"



	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�ͺţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_ggxh' type='text'>  <span class='tips'>����ո���ʾ�����Ѵ�������</span></td></tr>"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_ggxh"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_ggxh&btext=sb&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	
	
	dim sb_tablename,sb_tablebody,sb_table
			'ȡ�ֶε�����
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		'Dwt.out "<p align=""center"">��������</p>" 
	else
		do while not rsbody1.eof
			'�ֶ���
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'�ֶ���ҳ������ʾ������
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
		rsbody1.movenext
		loop
		sb_tablename=left(sb_tablename,len(sb_tablename)-1)  'ȥ�����ұ߶���
		sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  'ȥ�����ұ߶���
		sb_tablename=split(sb_tablename,",")
		sb_tablebody=split(sb_tablebody,",")
		for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
			dim sbtablename
			sbtablename=sb_tablename(sb_tablei)
			Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"�� </strong></td>"   & vbCrLf   
			Dwt.out"<td width='60%' class='tdbg'><input name='"&sbtablename&"' type='text'></td></tr>"& vbCrLf
		next
	end if 
	set rsbody1=nothing	

	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�������ң� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_sccj' type='text'>  <span class='red'>*</span><span class='tips'>����ո���ʾ�����Ѵ�������</span></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��Ʒ��ţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bh' type='text' ></td></tr>"& vbCrLf
   	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_sccj"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_sccj&btext=sb&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"

   
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺 </strong></td>"   & vbCrLf   
   Dwt.out"<td width='88%' class='tdbg'>"
   Dwt.out"<input name='sb_qydate' type='text'  onClick='new Calendar(0).show(this)' readOnly  value='"&date()&"'>"
    Dwt.out"</td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��ע�� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bz' type='text'></td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>ͼƬ�� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_img' id='sb_img' type='hidden' >"& vbCrLf
    Dwt.out "<iframe src='neweditor/editor.htm?id=sb_img&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
    dwt.out "</td></tr>"
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveaddsb'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>     <input  type='submit' name='Submit' value=' ��   �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	Dwt.out"</table></form>"

end sub

sub saveaddsb()

'��������
	sb_classid=request("sbclassid")
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sb"
	rsadd.open sqladd,conn,1,3
	rsadd.addnew
    on error resume next
    rsadd("sb_dclass")=ReplaceBadChar(Trim(Request("sbclassid")))
	rsadd("sb_sscj")=ReplaceBadChar(Trim(Request("sb_sscj")))
	rsadd("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsadd("sb_dclass")) then 	rsadd("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '�ж��Ƿ����ӷ���,�ٱ���
	rsadd("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsadd("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsadd("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsadd("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsadd("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsadd("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
	rsadd("sb_bh")=ReplaceBadChar(Trim(request("sb_bh")))
	rsadd("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	rsadd("sb_img")=Trim(request("sb_img"))
	
	
	    sb_isls=request("sb_isls")
	if sb_isls="on" then
	  sb_isls=true
	else
	  sb_isls=false
	end if  
	rsadd("sb_isls")=sb_isls
    
	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsadd("sb_iszj")=sb_iszj
    
	sb_isbw=request("sb_isbw")
	if sb_isbw="on" then
	  sb_isbw=true
	else
	  sb_isbw=false
	end if  
	rsadd("sb_isbw")=sb_isbw
    
	sb_isjl=request("sb_isjl")
	if sb_isjl="on" then
	  sb_isjl=true
	else
	  sb_isjl=false
	end if  
	rsadd("sb_isjl")=sb_isjl

	
	dim sb_tablename,sb_tablebody,sb_table
			'ȡ�ֶε�����
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		'Dwt.out "<p align=""center"">��������</p>" 
	else
		do while not rsbody1.eof
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	

	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  'ȥ�����ұ߶���
	sb_tablename=split(sb_tablename,",")
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
        rsadd(sbtablename)=ReplaceBadChar(Trim(request(sbtablename)))
	next
	
	rsadd("sb_sccj")=ReplaceBadChar(Trim(request("sb_sccj")))
	rsadd("sb_bz")=ReplaceBadChar(Trim(request("sb_bz")))
	rsadd("sb_update")=now()
	rsadd.update
	rsadd.close
	  Dwt.savesl "�豸����-"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&ReplaceBadChar(Trim(Request("sbclassid"))))(0),"���",ReplaceBadChar(Trim(Request("sb_wh")))
 	
	
	Dwt.out"<Script Language=Javascript>location.href='sb.asp?sbclassid="&sb_classid&"'</Script>"

end sub


sub saveedit()
'�༭����
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sb where sb_ID="&ReplaceBadChar(Trim(request("ID")))
	rsedit.open sqledit,conn,1,3
	on error resume next

	rsedit("sb_ssgh")=ReplaceBadChar(Trim(Request("sb_ssgh")))
	if zclassor(rsedit("sb_dclass")) then 	rsedit("sb_zclass")=ReplaceBadChar(Trim(Request("sb_zclass")))  '�ж��Ƿ����ӷ���,�ٱ���
	rsedit("sb_wh")=ReplaceBadChar(Trim(Request("sb_wh")))
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))
	rsedit("sb_zqqk")=ReplaceBadChar(Trim(Request("sb_zqqk")))	
	rsedit("sb_tyqk")=ReplaceBadChar(Trim(Request("sb_tyqk")))
	rsedit("sb_whqk")=ReplaceBadChar(Trim(Request("sb_whqk")))	
	rsedit("sb_fj")=ReplaceBadChar(Trim(Request("sb_fj")))
	rsedit("sb_ggxh")=ReplaceBadChar(Trim(request("sb_ggxh")))
	rsedit("sb_bh")=ReplaceBadChar(Trim(request("sb_bh")))
    rsedit("sb_qydate")=ReplaceBadChar(Trim(request("sb_qydate")))
	sb_isls=request("sb_isls")
	if sb_isls="on" then
	  sb_isls=true
	else
	  sb_isls=false
	end if  
	rsedit("sb_isls")=sb_isls
    
	sb_iszj=request("sb_iszj")
	if sb_iszj="on" then
	  sb_iszj=true
	else
	  sb_iszj=false
	end if  
	rsedit("sb_iszj")=sb_iszj
    
	sb_isbw=request("sb_isbw")
	if sb_isbw="on" then
	  sb_isbw=true
	else
	  sb_isbw=false
	end if  
	rsedit("sb_isbw")=sb_isbw
    
	sb_isjl=request("sb_isjl")
	if sb_isjl="on" then
	  sb_isjl=true
	else
	  sb_isjl=false
	end if  
	rsedit("sb_isjl")=sb_isjl

	dim sb_tablename,sb_tablebody,sb_table
			'ȡ�ֶε�����
	sqlbody1="SELECT sbtable_name from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		Dwt.out "<p align=""center"">��������</p>" 
	else
		do while not rsbody1.eof
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
		rsbody1.movenext
		loop
	end if 
	set rsbody1=nothing	
	sb_tablename=left(sb_tablename,len(sb_tablename)-1)  'ȥ�����ұ߶���
	sb_tablename=split(sb_tablename,",")
	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
        rsedit(sbtablename)=ReplaceBadChar(Trim(request(sbtablename)))
	next
	
	rsedit("sb_img")=Trim(request("sb_img"))
	rsedit("sb_sccj")=ReplaceBadChar(Trim(request("sb_sccj")))
	rsedit("sb_bz")=ReplaceBadChar(Trim(request("sb_bz")))
	rsedit("sb_update")=now()
	rsedit.update
	rsedit.close
	  Dwt.savesl "�豸����-"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&ReplaceBadChar(Trim(Request("sbclassid"))))(0),"�༭",ReplaceBadChar(Trim(Request("sb_wh")))
	Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub

sub edit()
	Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function checkadd(){" & vbCrLf
Dwt.out " if(document.form.sb_sscj.value==''){" & vbCrLf
Dwt.out "      alert('��ѡ���������䣡');" & vbCrLf
Dwt.out "   document.form.sb_sscj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_ssgh.value==0){" & vbCrLf
Dwt.out "      alert('��ѡ������װ�ã�');" & vbCrLf
Dwt.out "   document.form.sb_ssgh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf
Dwt.out " if(document.form.sb_wh.value==''){" & vbCrLf
Dwt.out "      alert('����дλ�ţ�');" & vbCrLf
Dwt.out "   document.form.sb_wh.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf


Dwt.out " if(document.form.sb_sccj.value==''){" & vbCrLf
Dwt.out "      alert('����д�������ң�');" & vbCrLf
Dwt.out "   document.form.sb_sccj.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
	sb_id=ReplaceBadChar(Trim(request("id")))

	sqledit="SELECT * from sb where sb_id="&sb_id
	set rsedit=server.createobject("adodb.recordset")
	rsedit.open sqledit,conn,1,1
	Dwt.out"<form method='post' action='sb.asp'  name='form' onsubmit='javascript:return checkadd();'>"
	Dwt.out"<table width='90%' border='0' align='center' cellpadding='2' cellspacing='1' class='border' >"
	Dwt.out"<tr class='title'>"& vbCrLf
	Dwt.out"<td height='22' colspan='2'><Div align=center><strong>"&sb_classname&"�༭ "
	Dwt.out"λ��:"& vbCrLf
	Dwt.out rsedit("sb_wh")&"</strong></Div></tr>"& vbCrLf
	
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�������䣺 </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_sscj'  disabled='disabled'  type='text' value='"&sscjh(rsedit("sb_sscj"))&"'></td></tr>"& vbCrLf
    Dwt.out"<input name='sb_sscj' type='hidden' value="&rsedit("sb_sscj")&">"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>����װ�ã� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_ssgh' size='1' >"
	call formgh (rsedit("sb_ssgh"),rsedit("sb_sscj"))
	Dwt.out"</select></td></tr>"
	
	
	if zclassor(rsedit("sb_dclass")) then
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>���ͣ� </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'><select name='sb_zclass' size='1' >"
		formzclass sb_classid,rsedit("sb_zclass") 
		Dwt.out"</select></td></tr>"& vbcrlf
    end if 
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>λ�ţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_wh' type='text' value='"&rsedit("sb_wh")&"'></td></tr>"& vbCrLf
	
'	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�豸���ԣ� </strong></td>"   & vbCrLf   
'	Dwt.out"<td width='60%' class='tdbg'>"
'	Dwt.out" <label><input type='checkbox' name='sb_isls' "
'	if rsedit("sb_isls") then Dwt.out "checked='checked'"
'	Dwt.out" />�Ƿ����� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_iszj' "
'	if rsedit("sb_iszj") then Dwt.out "checked='checked'"
'	Dwt.out" />�Ƿ��ܼ� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_isbw' "
'	if rsedit("sb_isbw") then Dwt.out "checked='checked'"
'	Dwt.out" />�Ƿ��� </label>"
'	Dwt.out" <label><input type='checkbox' name='sb_isjl' "
'	if rsedit("sb_isjl") then Dwt.out "checked='checked'"
'	Dwt.out" />�Ƿ�������� </label>"
'	
'	Dwt.out "</td></tr>"& vbCrLf


		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��ã� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_whqk' size='1' >"
	Dwt.out"<option value='0'"
	
	if rsedit("sb_whqk")=0 then Dwt.out" selected" 
	Dwt.out">��ѡ��������</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_whqk")=1 then Dwt.out"selected"
	Dwt.out">���</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_whqk")=2 then Dwt.out"selected"
	Dwt.out" >�����</option>"
	Dwt.out"</select></td></tr>"
	
	if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&sb_classid)(0) then 
		Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>׼ȷ�� </strong></td>"   & vbCrLf   
		Dwt.out"<td width='60%' class='tdbg'>"
		Dwt.out"<select name='sb_zqqk' size='1' >"
		Dwt.out"<option value='0'"
		if rsedit("sb_zqqk")=0 then Dwt.out" selected" 
		Dwt.out">��ѡ��׼ȷ���</option>"
		Dwt.out"<option value='1' "
		if rsedit("sb_zqqk")=1 then Dwt.out"selected"
		Dwt.out">�����С</option>"
		Dwt.out"<option value='2'"
		if rsedit("sb_zqqk")=2 then Dwt.out"selected"
		Dwt.out" >�м�</option>"
		Dwt.out"<option value='3'"
		if rsedit("sb_zqqk")=3 then Dwt.out"selected"
		Dwt.out" >>95%</option>"
		Dwt.out"</select></td></tr>"
    end if 

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>Ͷ�ˣ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_tyqk' size='1' >"
	Dwt.out"<option value='0'"
	if rsedit("sb_tyqk")=0 then Dwt.out" selected" 
	Dwt.out">��ѡ��Ͷ�����</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_tyqk")=1 then Dwt.out"selected"
	Dwt.out">Ͷ��</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_tyqk")=2 then Dwt.out"selected"
	Dwt.out" >ԭ��δͶ��</option>"
	Dwt.out"<option value='3' "
	if rsedit("sb_tyqk")=3 then Dwt.out"selected"
	Dwt.out">����ԭ��δͶ��</option>"
	Dwt.out"</select></td></tr>"

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�ּ��� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'>"
	Dwt.out"<select name='sb_fj' size='1' >"
	Dwt.out"<option value='0'"
	if rsedit("sb_fj")=0 then Dwt.out" selected" 
	Dwt.out">��ѡ��ּ�</option>"
	Dwt.out"<option value='1' "
	if rsedit("sb_fj")=1 then Dwt.out"selected"
	Dwt.out">��</option>"
	Dwt.out"<option value='2'"
	if rsedit("sb_fj")=2 then Dwt.out"selected"
	Dwt.out" >���</option>"
	Dwt.out"<option value='3' "
	if rsedit("sb_fj")=3 then Dwt.out"selected"
	Dwt.out">����</option>"
	Dwt.out"</select></td></tr>"
	
	Dwt.out "<link rel=""stylesheet"" type=""text/css"" href=""css/autocomplete.css"" /> "
    Dwt.out "<script type=""text/javascript"" src=""js/prototype.js""></script>"
    Dwt.out "<script type=""text/javascript"" src=""js/autocomplete.js""></script>"
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�ͺţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_ggxh' type='text' value='"&rsedit("sb_ggxh")&"'>  <span class='tips'>����ո���ʾ�����Ѵ�������</span></td></tr>"& vbCrLf
	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_ggxh"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_ggxh&btext=sb&orderby=sb_ggxh&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	dim sb_tablename,sb_tablebody,sb_table
			'ȡ�ֶε�����
	sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
	set rsbody1=server.createobject("adodb.recordset")
	rsbody1.open sqlbody1,conn,1,1
	if rsbody1.eof and rsbody1.bof then 
		 
	else
		do while not rsbody1.eof
			'Dwt.out "<td style=""border-bottom-style: solid;border-width:1px""><Div align=""center""><strong>"&rsbody1("sbtable_body")&"</strong></Div></td>"
			'�ֶ���
			sb_tablename=sb_tablename&rsbody1("sbtable_name")&","
			'�ֶ���ҳ������ʾ������
			sb_tablebody=sb_tablebody&rsbody1("sbtable_body")&","
			
		rsbody1.movenext
		loop
sb_tablename=left(sb_tablename,len(sb_tablename)-1)  'ȥ�����ұ߶���
	sb_tablebody=left(sb_tablebody,len(sb_tablebody)-1)  'ȥ�����ұ߶���
	sb_tablename=split(sb_tablename,",")
	sb_tablebody=split(sb_tablebody,",")


	for sb_tablei=LBound(sb_tablename) To UBound(sb_tablename) 
	   	dim sbtablename
		sbtablename=sb_tablename(sb_tablei)
		
'		��ɾ��if mid(sbtablename,4,1)="b" then
'		
'		'BOOL�����ֶ�,�Ե�һ����Ϊ��,�ڶ�����Ϊ��,��"��������" ����,���
'			Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"�� </strong></td>"   & vbCrLf   
'			Dwt.out"<td width='60%' class='tdbg'>"
'			Dwt.out"<select name='sbtablename' size='1' >"
'			Dwt.out"<option value='0'"
'			if rsedit(sbtablename)=0 then Dwt.out" selected" 
'			Dwt.out">��ѡ��"&sb_tablebody(sb_tablei)&"</option>"
'			Dwt.out"<option value='true' "
'			if rsedit(sbtablename)=true then Dwt.out"selected"
'			Dwt.out">"&left(sb_tablebody(sb_tablei),1)&"</option>"
'			Dwt.out"<option value='false'"
'			if rsedit(sbtablename)=false then Dwt.out"selected"
'			Dwt.out" >"&mid(sb_tablebody(sb_tablei),2,1)&"</option>"
'			Dwt.out"</select></td></tr>"
'		else 
			Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>"&sb_tablebody(sb_tablei)&"�� </strong></td>"   & vbCrLf   
			Dwt.out"<td width='60%' class='tdbg'><input name='"&sbtablename&"' type='text' value='"&rsedit(sbtablename)&"'></td></tr>"& vbCrLf
	   'end if 
		'Dwt.out sbtablename&"<br>"&sb_tablei
   'Dwt.out sb_tablename(sb_tablei)&" "&sb_tablebody(sb_tablei)
	next
	end if 
	set rsbody1=nothing	

	
	

	

	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>�������ң� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_sccj' type='text' value='"&rsedit("sb_sccj")&"'>  <span class='red'>*</span><span class='tips'>����ո���ʾ�����Ѵ�������</span></td></tr>"& vbCrLf
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��Ʒ��ţ� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bh' type='text' value='"&rsedit("sb_bh")&"'>  <span class='tips'>����ո���ʾ�����Ѵ�������</span></td></tr>"& vbCrLf
   	Dwt.out "<script>"
    '�Զ���ɺ��������Ϊѡ��״̬
    Dwt.out "function setSelectionRange(input, selectionStart, selectionEnd){if (input.setSelectionRange){input.setSelectionRange(selectionStart, selectionEnd);}else if (input.createTextRange) {var range = input.createTextRange();range.collapse(true);range.moveEnd('character', selectionEnd);range.moveStart('character', selectionStart);range.select();$(""wiki3"").focus();}}"
	Dwt.out "function update(object, value) {object.text.value = value;	var index = value.toLowerCase().indexOf(object.value.toLowerCase());if (index > -1) {setSelectionRange(object.text, index + object.value.length, value.length);}}"
	Dwt.out "new CAPXOUS.AutoComplete(""sb_sccj"", function() {return ""/inc/autocomplete.asp?dbname=data&zdtext=sb_sccj&btext=sb&typing="" + escape(this.text.value);});"
    Dwt.out "</script>"
	
	
	
   Dwt.out"<tr class='tdbg'><td width='12%' align='right' class='tdbg'><strong>����ʱ�䣺 </strong></td>"   & vbCrLf   
   Dwt.out"<td width='88%' class='tdbg'>"
   Dwt.out"<input name='sb_qydate' type='text' onClick='new Calendar(0).show(this)' readOnly  value="&rsedit("sb_qydate")&">"
   Dwt.out"</td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>��ע�� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input name='sb_bz' type='text' value='"&rsedit("sb_bz")&"'></td></tr>"& vbCrLf
	
	Dwt.out"<tr class='tdbg'><td width='40%' align='right' class='tdbg'><strong>ͼƬ�� </strong></td>"   & vbCrLf   
	Dwt.out"<td width='60%' class='tdbg'><input type='hidden' name='sb_img' id='sb_img' value='"&rsedit("sb_img")&"'>"& vbCrLf
    Dwt.out "<iframe src='neweditor/editor.htm?id=sb_img&ReadCookie=0' frameBorder='0' marginHeight='0' marginWidth='0' scrolling='No' width='621' height='457'></iframe>"
	dwt.out "</td></tr>"& vbCrLf
	Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'><input name='sbclassid' type='hidden' id='action' value='"&sb_classid&"'>   <input name='id' type='hidden'  value='"&Trim(Request("id"))&"'> <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="";history.back()"" style='cursor:hand;'></td>  </tr>"
	
	Dwt.out"</table></form>"
  rsedit.close
  set rsedit=nothing
  conn.close
  set conn=nothing

end sub


sub main()
 url= GetUrl
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function isDel(id){" & vbCrLf
Dwt.out "		if(confirm('��ȷ��Ҫɾ����������')){" & vbCrLf
Dwt.out "			location.href='sb.asp?action=del&id='+id;" & vbCrLf
Dwt.out "		}" & vbCrLf
Dwt.out "	}" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf

'	sqlbody="SELECT * from sb where sb_dclass="&sb_classid
'20111122����ʾ����վ����
	sqlbody="SELECT * from sb where sb_isdel=false and sb_dclass="&sb_classid

	if sscjid<>"" then sqlbody=sqlbody&" and sb_sscj="&sscjid
	if ssghid<>"" then sqlbody=sqlbody&" and sb_ssgh="&ssghid
	if keys<>"" then sqlbody=sqlbody&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlbody=sqlbody&" and sb_zclass="&request("sbzclassid")
	if request("update")<>"" then 
    	sqlbody=sqlbody&" order by sb_update desc"
	else
    	sqlbody=sqlbody&" order by sb_sscj aSC,sb_ssgh asc,sb_wh asc"
	end if 
        '���������ʵ�֣��������������У�δʵ��
    	'sqlbody=sqlbody&" order by [select * form sbgh where sb_id=sb.sb_id] desc"

	set rsbody=server.createobject("adodb.recordset")
	rsbody.open sqlbody,conn,1,1

	if request("sscj")<>"" then title=sscjh(sscjid)&"��" 
	if request("ssgh")<>"" then title=gh(ssghid) &"��"
	if request("keyword")<>"" then title=" '"&keys&" '"&"��"
    title="����"&title&sb_classname
	if request("sbzclassid")<>"" then title=title&"��"&conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&request("sbzclassid"))(0)
	
	
	Dwt.out "<Div style='left:6px;'>"
	Dwt.out "     <Div class='x-layout-panel-hd x-layout-title-center'>"
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>�豸����"&title&"</span>"
	Dwt.out "     </Div>"
'20111122����ʾ����վ����
	   'sqlcj="SELECT distinct sb_sscj from sb where sb_dclass="&sb_classid
        sqlcj="SELECT distinct sb_sscj from sb where  sb_isdel=false and sb_dclass="&sb_classid
		
		   sqlcj=sqlcj&" order by sb_sscj asc"
	   set rscj=server.createobject("adodb.recordset")
               rscj.open sqlcj,conn,1,1
               do while not rscj.eof
	       sscji=cint(rscj("sb_sscj"))
           ' for sscji=1 to 5 
	  sql="SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid&" and sb_sscj="&sscji
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	  sb_numb=sb_numb&sscjh_d(sscji)&":"&"<font color='#006600'>"&conn.Execute(sql)(0)&"</font>&nbsp;&nbsp;&nbsp;&nbsp;"
	   ' next
              rscj.movenext
	      loop
	      rscj.close
	      set rscj=nothing

	sql="SELECT count(sb_id) FROM sb WHERE sb_dclass="&sb_classid
	  if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
	totall= "<font color='#006600'>"&conn.Execute(sql)(0)&"</font>" 
	'Dwt.out "<Div class='pre'> <strong>άһ��"&v1&"</strong>   <strong>ά����"&v2&"</strong>     <strong>ά����"&v3&"</strong>     <strong>ά�ģ�"&v4&"</strong>     <strong>�ۺϣ�"&zh&"</strong>     <strong>�ϼƣ�"&totall&"</strong></Div>"
	Dwt.out "<Div class='pre'>"&sb_numb&"�ϼ�:"&totall&"<br/>λ��ǰ��<span style=""color:#0033ff"">��</span>��ʾ������¹�&nbsp;&nbsp;���<span style=""color:#006600"">��</span>�����<span style=""color:#ff0000"">��</span> &nbsp;&nbsp;Ͷ��<span style=""color:#006600"">��</span>��δͶ��<span style=""color:#0000ff"">��</span>����δͶ��<span style=""color:#ff0000"">��</span></Div>"& vbCrLf
	Dwt.out "<Div class='x-layout-container' style='top:0px;WIDTH: 815px; POSITION: relative; HEIGHT: 543px'>"& vbCrLf
	Dwt.out "<Div class='x-layout-panel x-layout-panel-center' style='LEFT: 3px; WIDTH: 810px; TOP: 3px; HEIGHT: 537px'>"& vbCrLf
	search	()
	
	if rsbody.eof and rsbody.bof then 
		message "<p align=""center"">δ�ҵ��������</p>" & vbCrLf
	else
	    Dwt.out "<SCRIPT src='js/grid.js' type=text/javascript></SCRIPT>"& vbCrLf
	    record=rsbody.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rsbody.PageSize = Cint(PgSz) 
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
		rsbody.absolutePage = page
		start=PgSz*Page-PgSz+1
		rowCount = rsbody.PageSize
		
		Dwt.out "<SCRIPT language=JavaScript >"& vbCrLf
        'Dwt.out "// ��λ���� ( ��λ���� # ��λ��� # ���϶��� )"
		Dwt.out "var DataTitles=new Array("& vbCrLf
		Dwt.out """���#40#center"","& vbCrLf
		Dwt.out """λ��#160#left"","& vbCrLf
		Dwt.out """����#120#center""  ,"& vbCrLf
		Dwt.out """װ��#90#center"","& vbCrLf
		if zclassor(rsbody("sb_dclass")) then Dwt.out """����#80 #center"","& vbCrLf
		Dwt.out """���#58 #center"","& vbCrLf
		
		'����ڷ������趨����ʾ"׼ȷ"����ʾ
		if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&rsbody("sb_dclass"))(0) then Dwt.out """׼ȷ#58 #center"","& vbCrLf
		
		Dwt.out """Ͷ��#58 #center"","& vbCrLf
		Dwt.out """�ּ�#58 #center"","& vbCrLf
		Dwt.out """�ͺ�#150#left"","& vbCrLf

		'ȡ�ֶε�����
		sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
		set rsbody1=server.createobject("adodb.recordset")
		rsbody1.open sqlbody1,conn,1,1
		if rsbody1.eof and rsbody1.bof then 
			'Dwt.out "<p align=""center"">��������</p>" 
		else
			do while not rsbody1.eof
				Dwt.out """"&rsbody1("sbtable_body")&"#140 #center"","& vbCrLf
				rsbody1.movenext
			loop
		end if 
		set rsbody1=nothing	
		Dwt.out """��������#80 #left"","& vbCrLf
		Dwt.out """��Ʒ���#150#left"","& vbCrLf
		Dwt.out """����ʱ��#70 #center"","& vbCrLf
		Dwt.out """��ע#100 #left"","& vbCrLf
		Dwt.out """ѡ��#80 #center"")</SCRIPT>"
		Dwt.out "<SCRIPT language=JavaScript >"
		Dwt.out "var DataFields=new Array()"& vbCrLf
		i=0
		do while not rsbody.eof and rowcount>0
				xh_id=((page-1)*pgsz)+1+xh
				xh=xh+1
			Dwt.out "DataFields["&i&"] =new Array("
			Dwt.out "'"&xh_id&"',"
			
			Dwt.out "'"
			if now()-rsbody("sb_update")<7 then Dwt.out "<span style=""color:#0033ff"">��</span>"
			Dwt.out searchH(uCase(rsbody("sb_wh")),keys)&"',"
			Dwt.out "'<a href=sb_jxjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&sb_classid&">��</a>&nbsp;<a href=sb_ghjl.asp?sbid="&rsbody("sb_id")&"&sbclassid="&sb_classid&">��</a>&nbsp;"
	        if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,rsbody("sb_sscj")) then Dwt.out "<a href=""sb.asp?action=edit&sbclassid="&sb_classid&"&id="&rsbody("sb_id")&""">��</a>&nbsp;"
			if conn.Execute("SELECT sb_img FROM sb WHERE sb_id="&rsbody("sb_id"))(0)<>"" then dwt.out "<a href=sb.asp?action=img&sbid="&rsbody("sb_id")&"  target=""_blank"">ͼ</a>&nbsp;"
			Dwt.out sscjh_d(rsbody("sb_sscj"))&"',"
			
			Dwt.out "'"&GH(rsbody("sb_ssGH"))&"',"
			if zclassor(rsbody("sb_dclass")) then 
			   if zclass(rsbody("sb_zclass"))="δ�༭" then 
			    dwt.out  "'"&zclass(rsbody("sb_dclass"))&"',"
			   else
			    Dwt.out "'"&zclass(rsbody("sb_zclass"))&"'," 
			   end if 
			 end if   	
			'Dwt.out """"&xh_id&""","
			Dwt.out "'"&sb_whd(rsbody("sb_whqk"))&"',"
			if conn.Execute("SELECT sbclass_zq FROM sbclass WHERE sbclass_id="&rsbody("sb_dclass"))(0) then Dwt.out "'"&sb_zqd(rsbody("sb_ZQqk"))&"',"
			Dwt.out "'"&sb_tyd(rsbody("sb_tyqk"))&"',"
			Dwt.out "'"&fj(rsbody("sb_fj"))&"',"
			Dwt.out "'"&rsbody("sb_ggxh")&"',"

		
			'ȡ�ֶ�����
			sqlbody1="SELECT * from sbtable where sbtable_sbclassid="&sb_classid&" order by  sbtable_orderby aSC"
			set rsbody1=server.createobject("adodb.recordset")
			rsbody1.open sqlbody1,conn,1,1
			if rsbody1.eof and rsbody1.bof then 
				'Dwt.out "<p align=""center"">��������</p>" 
			else
				do while not rsbody1.eof
				  sbtable_name=rsbody1("sbtable_name")   'ȡ�ñ������
				  Dwt.out "'"&rsbody(sbtable_name)&"',"
				  'message sbtable_name
				rsbody1.movenext
				loop
			end if 
			set rsbody1=nothing	

			Dwt.out "'"&rsbody("sb_sccj")&"',"
			Dwt.out "'"&rsbody("sb_bh")&"',"
			Dwt.out "'"&rsbody("sb_qydate")&"',"
			Dwt.out "'"&rsbody("sb_bz")&"',"
			Dwt.out "'"
			call sbeditdel(rsbody("sb_id"),rsbody("sb_sscj"),"sb.asp?action=edit&sbclassid="&sb_classid&"&id=")'���ޡ��������༭��ɾ��
			Dwt.out "')"& vbCrLf
			
			RowCount=RowCount-1
					i=i+1
		rsbody.movenext
		loop
		Dwt.out "</script>"
	Dwt.out "<TABLE cellSpacing=0 cellPadding=0 border=0>"
	Dwt.out "  <TBODY>"
	Dwt.out "  <tr>"
	Dwt.out "    <TD valign='top' style='BORDER-RIGHT: white 2px inset; BORDER-TOP: white 2px inset; BORDER-LEFT: white 2px inset; BORDER-BOTTOM: white 2px inset; BACKGROUND-COLOR: scrollbar'>"
	Dwt.out "      <Div id=DataTable></Div></TD></tr></TBODY></TABLE>"
		call sbshowpage(page,url,total,record)
	Dwt.out "</Div></Div></Div>"
	end if
	rsbody.close
	set rsbody=nothing
	conn.close
	set conn=nothing

end sub
'	Dwt.out "����ִ����ʱ��" & timer-starttime

Dwt.out "</body></html>"

'ѡ��༭��ɾ����
sub sbeditdel(id,sscj,editurl)
'	if displaypagelevelh(session("groupid"),2,session("pagelevelid")) and displaygrouplevelh(session("groupid"),0,sscj) then 
'	
'	Dwt.out "<a href="""&editurl&id&""">�༭</a>&nbsp;"
'	end if 	
	if  displaypagelevelh(session("groupid"),3,session("pagelevelid")) and displaygrouplevelh(session("groupid"),1,sscj)  then
	 Dwt.out "<a href=""javascript:isDel("&id&");"">ɾ��</a>"
	end if 
	Dwt.out "&nbsp;"
end sub



'ȡ�ӷ�������
function zclass(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_id="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclass= "δ�༭"
  else
     zclass=rsbody("sbclass_name")
  end if
end function

'�ж��Ƿ����ӷ���
function zclassor(id)
dim sqlbody,rsbody
 sqlbody="SELECT * from sbclass where sbclass_zclass="&id
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     zclassor=false 
  else
     zclassor=true
  end if
end function


'�������б���ʾ
function formdclass()
	dim sqldclass,rsdclass
	'if isnull(dclassid) then dclassid=0
'	if dclassid=0 then 
		sqldclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_isput=true"
'	else
'		sqldclass="SELECT * from sbclass where sbclass_dclass<>0 and sbclass_id="&dclassid
'	end if 		
	set rsdclass=server.createobject("adodb.recordset")
	rsdclass.open sqldclass,conn,1,1
	if rsdclass.eof then 
		dclass="û���κη���"
	else
		Dwt.out"<option value='0'"
		if dclassid=0 then Dwt.out " selected" 
			Dwt.out">��ѡ��Ҫ����豸�ķ���</option>"
		do while not rsdclass.eof
			Dwt.out"<option value='sb.asp?action=addsb&sbclassid="&rsdclass("sbclass_id")&"'>"&rsdclass("sbclass_name")&"</option>"  & vbCrLf   
		rsdclass.movenext
		loop
	end if 
	rsdclass.close
	set rsdclass=nothing
end function


'�ӷ����б���ʾ
function formzclass(dclassid,zclassid)
	dim sqlzclass,rszclass
	if isnull(zclassid) then zclassid=0
'	if zclassid=0 then 
		sqlzclass="SELECT * from sbclass  where sbclass_zclass<>0 and sbclass_zclass="&dclassid
'	else
		'sqlzclass="SELECT * from sbclass where sbclass_zclass<>0 and sbclass_id="&zclassid
'	end if 		
	set rszclass=server.createobject("adodb.recordset")
	rszclass.open sqlzclass,conn,1,1
	if rszclass.eof then 
		formzclass="δ�༭"
	else
		Dwt.out"<option value='0'"
		if zclassid=0 then Dwt.out " selected" 
			Dwt.out">��ѡ������</option>"
		do while not rszclass.eof
			Dwt.out"<option value='"&rszclass("sbclass_id")&"' "
			if zclassid=rszclass("sbclass_id") then Dwt.out "selected"
			Dwt.out">"&rszclass("sbclass_name")&"</option>"  & vbCrLf   
		rszclass.movenext
		loop
	end if 
	rszclass.close
	set rszclass=nothing
end function

'��������ʾ
Function sb_whd(whnumb)
	if isnull(whnumb) or whnumb=0 then 
	  sb_whd="δ�༭"
	else
		if whnumb=1 then sb_whd="<span style=""color:#006600"">��</span>"  '�����
		if whnumb=2 then sb_whd="<span style=""color:#ff0000"">��</span> "	  '����ú�
	end if 
end Function 

'׼ȷ�����ʾ
Function sb_zqd(zqnumb)

	if isnull(zqnumb) or zqnumb=0 then 
	  sb_zqd="δ�༭"
	else
		if zqnumb=3 then sb_zqd="����"'>95%
		if zqnumb=2 then sb_zqd="���"		  '�м�  
		if zqnumb=1 then sb_zqd="��"  '�����С
	end if 
end Function 

'Ͷ�������ʾ
Function sb_tyd(tynumb)

	if isnull(tynumb) or tynumb=0 then 
	  sb_tyd="δ�༭"
	else
		if tynumb=1 then sb_tyd="<span style=""color:#006600"">��</span>"   '��Ͷ��
		if tynumb=2 then sb_tyd="<span style=""color:#0000ff"">��</span>"   '����ԭ��δͶ��
		if tynumb=3 then sb_tyd="<span style=""color:#ff0000"">��</span>"    '�칤��ԭ��δͶ��
		'if zqnumb=4 then sb_zqd="<font color='#ff0000'>*</font>"    '�칤��ԭ��δͶ��
	end if 
end Function 



sub search()
	dim rscj,sqlcj,sscjid
	Dwt.out"<script type=""text/javascript"" src=""js/function.js""></script>"&vbcrlf
	
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	
	Dwt.out "<form method='Get' name='SearchForm' action='sb.asp'>" & vbCrLf
	if displaypagelevelh(session("groupid"),1,session("pagelevelid")) then Dwt.out "<a href='sb.asp?action=addsb&sbclassid="&sb_classid&"'>����豸</a>"
	Dwt.out "&nbsp;&nbsp;<select   onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">" & vbCrLf
	Dwt.out "<option value=''>��ʾ˳��ѡ��</option>" & vbCrLf
	Dwt.out "<option value='sb.asp?update=update&sbclassid="&sb_classid&"'>������ʱ��</option>"
	Dwt.out "     </select>	" & vbCrLf

	
	Dwt.out "  <input type='hidden' name='sbclassid' value='"&sb_classid&"'>" & vbCrLf
	if request("sbzclassid")<>"" then Dwt.out "<input type='hidden' name='sbzclassid' value='"&request("sbzclassid")&"'>" & vbCrLf

	Dwt.out "  <input type='text' name='keyword' id=""keyword"" size='20' maxlength='50'"
	if request("keyword")<>"" then 
	 Dwt.out "value='"&request("keyword")&"'"
    	Dwt.out ">" & vbCrLf
    else
	 Dwt.out "value='����������λ��'"
	 	Dwt.out " onblur=""if(this.value==''){this.value='����������λ��'}"" onfocus=""this.value=''"">" & vbCrLf
	end if    
	Dwt.out "  <input type='submit' name='Submit'  value='����'>"
	



	
	Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tosscj(this.options[this.selectedIndex].value);}"">" & vbCrLf
	Dwt.out "<option value=''>��������ת����</option>" & vbCrLf
	sqlgh="SELECT distinct sb_sscj from sb where sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		cjid=cint(rsgh("sb_sscj"))


		sql="SELECT count(sb_id) FROM sb WHERE sb_sscj="&cjid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			'i=i+1
			Dwt.out"<option value='"&cjid&"'"
			if cint(request("sscj"))=cjid then Dwt.out" selected"
			sql="SELECT levelname FROM levelname WHERE levelid="&cjid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    cj_name="δ֪��"
			else
			    cj_name=rs("levelname")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&cj_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf

















'	sqlcj="SELECT * from levelname where levelclass=1 "& vbCrLf
'		set rscj=server.createobject("adodb.recordset")
'		rscj.open sqlcj,conn,1,1
'		do while not rscj.eof
'			Dwt.out"<option value='"&rscj("levelid")&"' "
'			if cint(request("sscj"))=rscj("levelid") then Dwt.out" selected"
'			Dwt.out ">"&rscj("levelname")&"</option>"& vbCrLf	
'			rscj.movenext
'		loop
'		rscj.close
'		set rscj=nothing
'		Dwt.out "     </select>	" & vbCrLf



'Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tossgh(this.options[this.selectedIndex].value);}"">" & vbCrLf
'Dwt.out "	       <option value=''>��װ����ת����</option>" & vbCrLf
'sscjid=session("levelclass")
'if sscjid=7 or sscjid=8 then
'sqlgh="SELECT * from ghname  ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
'else
'sqlgh="SELECT * from ghname where sscj="&sscjid&" ORDER BY SSCJ ASC,gh_name ASC"& vbCrLf
'end if
'    set rsgh=server.createobject("adodb.recordset")
'    rsgh.open sqlgh,conn,1,1
'    do while not rsgh.eof
'		sb_numb=Conn.Execute("SELECT count(sb_id) FROM sb WHERE sb_ssgh="&rsgh("ghid")&"and sb_dclass="&sb_classid)(0)
'		if sb_numb<>0 then 
'			i=i+1
'			Dwt.out"<option value='"&rsgh("ghid")&"'"
'			if cint(request("ssgh"))=rsgh("ghid") then Dwt.out" selected"
'			Dwt.out ">"&i&"&nbsp;&nbsp;"&rsgh("gh_name")&"("&sb_numb&")</option>"& vbCrLf
'	    end if 
'		rsgh.movenext
'	loop
'	rsgh.close
'	set rsgh=nothing
'	Dwt.out "     </select>	" & vbCrLf
	
	
	

	
	
Dwt.out "&nbsp;&nbsp;<select onchange=""if(this.options[this.selectedIndex].value!=''){tossgh(this.options[this.selectedIndex].value);}"">" & vbCrLf
Dwt.out "	       <option value=''>��װ����ת����</option>" & vbCrLf



	sqlgh="SELECT distinct sb_ssgh,sb_sscj from sb where sb_isdel=false and sb_dclass="&sb_classid
	if keys<>"" then sqlgh=sqlgh&" and sb_wh  like '%" &keys& "%' "
	if request("sbzclassid")<>"" then sqlgh=sqlgh&" and sb_zclass="&request("sbzclassid")
    sqlgh=sqlgh&" order by sb_sscj asc"
	set rsgh=server.createobject("adodb.recordset")
    rsgh.open sqlgh,conn,1,1
    do while not rsgh.eof
		ghid=cint(rsgh("sb_ssgh"))


		sql="SELECT count(sb_id) FROM sb WHERE sb_isdel=false and  sb_ssgh="&ghid&"and  sb_dclass="&sb_classid
		if keys<>"" then sql=sql&" and sb_wh  like '%" &keys& "%' "
		if request("sbzclassid")<>"" then sql=sql&" and sb_zclass="&request("sbzclassid")
        
		sb_numb=Conn.Execute(sql)(0)
        
		if sb_numb<>0 then 
			i=i+1
			Dwt.out"<option value='"&ghid&"'"
			if cint(request("ssgh"))=ghid and request("ssgh")<>"" then Dwt.out" selected"
			
			sql="SELECT gh_name FROM ghname WHERE ghid="&ghid
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			if rs.eof then 
			    gh_name="δ֪��"
			else
			    gh_name=rs("gh_name")
			end if 		
			rs.close
			set rs=nothing	
			Dwt.out ">"&i&"&nbsp;&nbsp;"&gh_name&"("&sb_numb&")</option>"& vbCrLf '
	    end if 
		rsgh.movenext
	loop
	rsgh.close
	set rsgh=nothing
	Dwt.out "     </select>	" & vbCrLf


	Dwt.out "</form></Div></Div>" & vbCrLf
	
	
end sub

'********************************************8
'��ҳ��ʾpage��ǰҳ����url��ҳ��ַ��total��ҳ�� record����Ŀ��
'pgsz ÿҳ��ʾ��Ŀ��
'URL�д�����
'*******************************************
sub sbshowpage(page,url,total,record)
   Dwt.Out "<Div class='x-toolbar'>"
   if page="" then page=1
   if page > 1 Then 
      Dwt.Out "<a href="&url&"&page=1><img src=images/top.gif border=0 align=absmiddle></a>&nbsp;<a href="&url&"&pgsz="&pgsz&"&page="&page-1&"><img src=images/page1.gif border=0 align=absmiddle></a>&nbsp;"
   else
      Dwt.Out ""
   end if 
   if RowCount = 0 and page <>Total then 
     Dwt.Out "<a href="&url&"&pgsz="&pgsz&"&page="&page+1&"><img src=images/page2.gif border=0 align=absmiddle></a> <a href="&url&"&pgsz="&pgsz&"&page="&total&"><img src=images/down.gif border=0 align=absmiddle></a>"
   else
     Dwt.Out ""
   end if
   Dwt.Out"&nbsp;&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&total&"</strong>ҳ&nbsp;&nbsp;"
  'if Total =1 then 
  '  Dwt.Out"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3'  disabled='disabled'  maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">��"
  'else
  ' Dwt.Out"ÿҳ��ʾ<input type='text' name='MaxPerPage' size='3' maxlength='4' value='"&pgsz&"' onKeyPress=""if (event.keyCode==13) window.location='"&url&"&pgsz='+this.value;"">��"
 ' end if 
   if Total =1 then 
    Dwt.Out"&nbsp;&nbsp;   <select name='1' disabled='disabled' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   else
    Dwt.Out"&nbsp;&nbsp;   <select name='1' onchange=""javascript:window.location='"&url&"&pgsz="&pgsz&"&page='+this.options[this.selectedIndex].value;"">"
   end if 
   for ii=1 to Total
     if ii=page then 
    	 Dwt.Out"  <option value='"&page&"' selected >��"&page&"ҳ</option>"
     else
    	 Dwt.Out"  <option value='"&ii&"'>��"&ii&"ҳ</option>"
     end if 
   next 
   
   Dwt.Out" </select>&nbsp;&nbsp;��"&record&"������"
   Dwt.Out "</Div>"
end sub

%>