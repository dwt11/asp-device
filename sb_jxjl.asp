<%@language=vbscript codepage=936 %>
<%
'Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/session.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/sb_function.asp"-->
<%dim url
dim sqlbody,rsbody,ii,sbclassid,ylbid,sqljx,rsjx
dim record,pgsz,total,page,rowCount,xh,sscj
dim sb_wh,sql,rs


sb_id=Trim(Request("sbid"))
sbclass_id=Trim(Request("sbclassid"))
url="sb_jxjl.asp?sbid="&sb_id&"&sbclassid="&sbclass_id
'��ȡ���࣬�����ڱ���
if sbclass_id="" or sb_id="" then Dwt.out"<Script Language=Javascript>history.back()</Script>"
sb_classname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE  sbclass_id="&sbclass_id)(0)
'if sb_id<>"" then 
 sb_wh=conn.Execute("SELECT sb_wh FROM sb WHERE  sb_id="&sb_id)(0)
 sb_sscj=conn.Execute("SELECT sb_sscj FROM sb WHERE  sb_id="&sb_id)(0)
'end if 
Dwt.out"<html>"& vbCrLf
Dwt.out"<head>" & vbCrLf
Dwt.out"<title> ������������ҳ</title>"& vbCrLf
Dwt.out"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbCrLf
Dwt.out "<link href='css/ext-all.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<link href='css/body.css' rel='stylesheet' type='text/css'>"& vbCrLf
Dwt.out "<SCRIPT language=javascript>" & vbCrLf
Dwt.out "function CheckAdd(){" & vbCrLf


%>
        var checkName = document.getElementsByName ("jx_gzxx_new");	//�����������ȡ�齨����
		var ischecked
		//ѭ��checkbox���ж��Ƿ����ѡ����
		for (i = 0; i < checkName.length; i ++) {
			if (checkName[i].checked) {	//�����ѡ����򷵻�true
				ischecked= true;
                
                if(checkName[i].value==0){
                  if(document.getElementById("jx_gzxx").value==''){
                         alert('����������Ϊ�գ�');
                      document.getElementById("jx_gzxx").focus();
                          return false;
                    }
                }
			}
		}
		if(!ischecked)
		 {alert('δѡ���������');
		  return false;	//ѭ����������ѡ�������FALSE
		 }

         checkName = document.getElementsByName ("jx_nr_new");	//�����������ȡ�齨����
		 ischecked=false
		//ѭ��checkbox���ж��Ƿ����ѡ����
		for (i = 0; i < checkName.length; i ++) {
			if (checkName[i].checked) {	//�����ѡ����򷵻�true
				ischecked= true;
                if(checkName[i].value==0){
                  if(document.getElementById("jx_nr").value==''){
                         alert('�������ݲ���Ϊ�գ�');
                      document.getElementById("jx_nr").focus();
                          return false;
                    }
                }
                if(checkName[i].value==9999){
                  if(document.getElementById("gh_xh").value==''){
                         alert('����ǰ�ͺŲ���Ϊ�գ�');
                      document.getElementById("gh_xh").focus();
                          return false;
                    }
                  if(document.getElementById("gh_xhupdate").value==''){
                         alert('�������ͺŲ���Ϊ�գ�');
                      document.getElementById("gh_xhupdate").focus();
                          return false;
                    }
                }
                
                
                
			}
		}
		if(!ischecked)
		 {alert('δѡ��������ݣ�');
		  return false;	//ѭ����������ѡ�������FALSE
		 }

<%




Dwt.out "  if(document.formadd.jx_fzren.value==''){" & vbCrLf
Dwt.out "      alert('���޸����˲���Ϊ�գ�');" & vbCrLf
Dwt.out "  document.formadd.jx_fzren.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf



Dwt.out "  if(document.formadd.jx_ren.value==''){" & vbCrLf
Dwt.out "      alert('�����˲���Ϊ�գ�');" & vbCrLf
Dwt.out "  document.formadd.jx_ren.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

Dwt.out "  if(document.formadd.jx_date.value==''){" & vbCrLf
Dwt.out "      alert('����ʱ�䲻��Ϊ�գ�');" & vbCrLf
Dwt.out "  document.formadd.jx_date.focus();" & vbCrLf
Dwt.out "      return false;" & vbCrLf
Dwt.out "    }" & vbCrLf

'Dwt.out "execScript('t=IsDate(document.formadd.jx_date.value)','VBScript');" & vbCrLf
'Dwt.out "if(!t){" & vbCrLf
'Dwt.out"   alert('���ڸ�ʽ����ȷ��Ӧ��yyyy-mm-dd');" & vbCrLf
'Dwt.out "  document.formadd.jx_date.focus();" & vbCrLf
'Dwt.out "return false;" & vbCrLf
'Dwt.out "    }" & vbCrLf


Dwt.out "    }" & vbCrLf
Dwt.out "</SCRIPT>" & vbCrLf
Dwt.out"<script language='javascript' type='text/javascript' src='js/My97DatePicker/WdatePicker.js'></script>"
Dwt.out"</head>"& vbCrLf
action=request("action")

Dwt.out"<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'"
'����Ǳ༭ ��JS���� ��ʱ���� ,�����͸����Ŀ��ж���ʾ
if action="edit" then dwt.out " onload='pickedFunc();clickgzxxqt();clickjxnrqt();clickjxnrgh()' "
'��������� ��JS���� ��ʱ����
if action="add" then dwt.out " onload='pickedFunc();' "
dwt.out ">"& vbCrLf

	Dwt.out "<Div style='left:6px;'>"& vbCrLf
	Dwt.out "     <Div class='x-layout-panel-hd'>"& vbCrLf
	Dwt.out "        <SPAN class='x-layout-panel-hd-text' style:'top:0px;'>"&sb_classname&"  "&sb_wh&" ���޼�¼</span>"& vbCrLf
	Dwt.out "     </Div>"& vbCrLf


select case action
  case "add"
      call add'����豸����ѡ��
  case "saveadd"
      call saveadd'����豸����ѡ��
  case "edit"
      call edit
  case "saveedit"'�༭�ӷ���
      call saveedit'�༭�����ӷ���
  case "del"
      call del     'ɾ���ӷ�����Ϣ
  case ""
      call main
end select	  	 




sub main()
	Dwt.out "<Div class='x-toolbar'><Div align=left>" & vbCrLf
	Dwt.out "<a href='sb.asp?sbclassid="&sbclass_id&"&keyword="&sb_wh&"'>����鿴 "&sb_wh&" ����ϸ��Ϣ</a> "

	sqljx="SELECT * from sbjx where sb_id="&sb_id&" order by  jx_DATE DESC"
	set rsjx=server.createobject("adodb.recordset")
	rsjx.open sqljx,conn,1,1
	if rsjx.eof and rsjx.bof then 
		if session("levelclass")=sb_sscj or session("levelclass")=0 then
			Dwt.out"<input type='button' name='Submit'  onclick=""window.location.href='sb_jxjl.asp?action=add&sbid="&sb_id&"&sbclassid="&sbclass_id&"'""value='��Ӽ��޼�¼'>"
		end if 	
		Dwt.out"<input name='Cancel' type='button' id='Cancel' value=' ��  �� ' onClick="";history.back()"" style='cursor:hand;'>"
		Dwt.out "</Div></Div>"
		message("δ���  "&sb_wh&" ���޼�¼")
	else
		if session("levelclass")=sb_sscj or session("levelclass")=0 then
			Dwt.out"<input type='button' name='Submit'  onclick=""window.location.href='sb_jxjl.asp?action=add&sbid="&sb_id&"&sbclassid="&sbclass_id&"'""value='��Ӽ��޼�¼'>"
		end if 	
		Dwt.out"<input name='Cancel' type='button' id='Cancel' value=' ��  �� ' onClick="";history.back()"" style='cursor:hand;'>"
		Dwt.out "</Div></Div>"
		
		record=rsjx.recordcount
		if Trim(Request("PgSz"))="" then
			PgSz=20
		ELSE 
			PgSz=Trim(Request("PgSz"))
		end if 
		rsjx.PageSize = Cint(PgSz) 
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
		rsjx.absolutePage = page
		dim start
		start=PgSz*Page-PgSz+1
		rowCount = rsjx.PageSize
		Dwt.out "<Div class='x-layout-panel' style='WIDTH: 100%;'>"& vbCrLf
		Dwt.out "<table width=""100%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"& vbCrLf
		Dwt.out "<tr class=""x-grid-header"">" & vbCrLf
		Dwt.out "     <td class='x-td'><Div class='x-grid-hd-text'></Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td' ><Div class='x-grid-hd-text'>�������</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ʼʱ��</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>����ʱ��</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ʱ(Сʱ)</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��������</Div></td>"& vbCrLf
		Dwt.out "      <td   class='x-td'><Div class='x-grid-hd-text'>��ע</Div></td>"& vbCrLf
		Dwt.out "      <td  class='x-td'><Div class='x-grid-hd-text'>ѡ��</Div></td>"& vbCrLf
		Dwt.out "    </tr>"& vbCrLf
		 do while not rsjx.eof and rowcount>0
			
			xh_id=((page-1)*pgsz)+1+xh
			xh=xh+1
			
			if xh_id mod 2 =1 then 
			  Dwt.out "<tr class='x-grid-row x-grid-row-alt' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			else
			  Dwt.out "<tr class='x-grid-row' onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#BFDFFF'"">"& vbCrLf
			end if 
			Dwt.out "     <td  style=""border-bottom-style: solid;border-width:1px"" ><Div align=""center"">"&xh_id&"</Div></td>"& vbCrLf
			
			
			jxlb=""
			if not isnull( rsjx("jx_ylwt") )  then jxlb="<span style=""color:#ff0000"">��</span> "  	  '����������Ϊ��

			
			
			if rsjx("jx_lb")<>"" then 
    			jxlb=jxlb&getjxlb(rsjx("jx_lb"))
            else
			    jxlb=jxlb&""
			end if 
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jxlb&"&nbsp;</td>"& vbCrLf
			
			
			
			
			
			jx_gzxx=""
			if not isnull( rsjx("jx_gzxx_new") ) then 
			  sbclassgz1=split(rsjx("jx_gzxx_new"),",")
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 then
							 '��ȡ���������� 
							 jxgzname=getjxgzxx(sbclassgz1(i))
							 'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +'��'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& jxgzname 
							else
							'��ȡ��������
										  if i<>0 then jx_gzxx=jx_gzxx & "<br>" 
										  jx_gzxx=jx_gzxx& "������"&rsjx("jx_gzxx") 
							end if 		  
							   		  
				   Next 
		    else
			    jx_gzxx="�����ݣ�"&rsjx("jx_gzxx") 
			end if  
			
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_gzxx&"&nbsp;</td>"& vbCrLf
			
			
			
			
			jx_nr=""
			if not isnull( rsjx("jx_nr_new") ) then 
			  sbclassgz1=split(rsjx("jx_nr_new"),",")
			   
			   
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							 if sbclassgz1(i)<>0 and sbclassgz1(i)<>9999 then     '0��������   99999�������
							 '��ȡ���������� 
								jxnrname=getjxnr(sbclassgz1(i))
								
										 'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +'��'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& jxnrname 
							else
							'��ȡ��������
								if sbclassgz1(i)=0 then 
										  if i<>0 then jx_nr=jx_nr & "<br>" 
										  jx_nr=jx_nr& "������"&rsjx("jx_nr")
								end if 
							'��ȡ��������
								if sbclassgz1(i)=9999 then 
										  if i<>0 then jx_nr=jx_nr & "<br>"
											'����Ҫ��� ��������Ϣ
										jx_nr=jx_nr& "������"
										sqlgh="SELECT gh_xh,gh_xhupdate FROM sbgh  WHERE jx_id="&rsjx("jx_id")
										set rsgh=server.createobject("adodb.recordset")
										rsgh.open sqlgh,conn,1,1
										if rsgh.eof and rsgh.bof then 
												jx_nr=jx_nr&"δ�ҵ��������ͺ�����"
										else
'											 if jx_nr= "" then 
'												jx_nr="����������ǰ�ͺ�<b>"&ghxh&"</b>���������ͺ�<B>"&ghxhupdate
'											 else
												jx_nr=jx_nr&"����ǰ�ͺ�<b>"&rsgh("gh_xh")&"</b>���������ͺ�<B>"&rsgh("gh_xhupdate")
'											 end if 
										end if   
		   
								end if 
							end if 			  
				   Next 
				   
				   'if jx_nr<> "" then 	       jx_nr=jx_nr&"<br>������"&rsjx("jx_nr") else  jx_nr="������"&rsjx("jx_nr") 
			else
			      ' (��ȡ�ɵ����ݻ���������
			      jx_nr="�����ݣ�"&rsjx("jx_nr")
			end if  
			

			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px;word-break:break-all;word-wrap:break-word"">"&jx_nr&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsjx("jx_date")&"</Div></td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"&rsjx("jx_enddate")&"</Div></td>"& vbCrLf
			
			a=rsjx("jx_date")
			b=rsjx("jx_enddate") 
			if not isnull(b) then 
			    ys=FormatNumber(DateDiff("n", a, b)/60,2,-1,0,0)
		     else
			    ys=""		
		     end if 		
			
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px""><Div align=""center"">"&ys&"</Div></td>"& vbCrLf
			Dwt.out "      <td  style=""cloudchen:expression(noWrap=true);border-bottom-style: solid;border-width:1px"">"&rsjx("jx_fzren")&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjx("jx_ren")&"&nbsp;</td>"& vbCrLf
			
			
			jx_ylwt=""
			if not isnull( rsjx("jx_ylwt") ) then 
			  sbclassgz1=split(rsjx("jx_ylwt"),",")
				   For i = LBound(sbclassgz1) To UBound(sbclassgz1)
							  
							 jxylwtname= getjxylwt(sbclassgz1(i))
							 ' conn.Execute("SELECT  sbjxylwt.sbjxylwt_name as sbjxylwt_name FROM sbjxylwt AS sbjxylwt left join sbjxylwt as sbjxylwtA on sbjxylwt.sbjxylwt_zclass=sbjxylwtA.sbjxylwt_id WHERE sbjxylwt.sbjxylwt_id="&sbclassgz1(i))(0)
										  if i<>0 then jx_ylwt=jx_ylwt & "<br>" 
										  jx_ylwt=jx_ylwt& jxylwtname 
				   Next 
		    else
			    		jx_ylwt="��"
			end if  
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&jx_ylwt&"&nbsp;</td>"& vbCrLf
			Dwt.out "      <td  style=""border-bottom-style: solid;border-width:1px"">"&rsjx("jx_bz")&"&nbsp;</td>"& vbCrLf
			Dwt.out" <td style=""border-bottom-style: solid;border-width:1px""><Div align=""center"">"
			call jxxuanxiang(rsjx("jx_id"),sb_id,sb_sscj)
			Dwt.out"</Div></td></tr>"			
			
			RowCount=RowCount-1
		rsjx.movenext
		loop
		Dwt.out"</table>"
		call showpage(page,url,total,record,PgSz)
	end if
	rsjx.close
	set rsjx=nothing
	conn.close
	set conn=nothing
end sub


sub add()
   '����λ�ż��޼�¼
   Dwt.out"<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'  >"& vbCrLf
   Dwt.out"<form method='post' action='sb_jxjl.asp' name='formadd' onsubmit='javascript:return CheckAdd();'>"& vbCrLf
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.out"<Div align='center'><strong>����   "&sb_wh&" ���޼�¼</strong></Div></td>    </tr>"& vbCrLf
  
  
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
  dim ischecked  '�����жϼ�������һ�� �Ƿ�Ĭ��ѡ��
  ischecked=false
   
    dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxlb where sbjxlb_zclass=0 order by  sbjxlb_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
	  do while not rsbody.eof 
						'����
				sqlz="SELECT * from sbjxlb where sbjxlb_zclass="&rsbody("sbjxlb_id")&" order by  sbjxlb_orderby aSC"& vbCrLf
				set rsz=server.createobject("adodb.recordset")
				rsz.open sqlz,conn,1,1
				if rsz.eof and rsz.bof then 
					dwt.out"<input type='radio' name='jx_lb' value='"&rsbody("sbjxlb_id")&"'" 
							if not ischecked then 
							   ischecked=true
							   dwt.out "checked"
							end if    
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name") & "<br>"
				else
					do while not rsz.eof
					
						dwt.out"<input type='radio' name='jx_lb' value='"&rsz("sbjxlb_id")&"'" 
							if not ischecked then 
							   ischecked=true
							   dwt.out "checked"
							end if    
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name")&":"&rsz("sbjxlb_name") & "<br>"
					rsz.movenext
					loop
				end if 	
				rsz.close
				set rsz=nothing
			
		rsbody.movenext
		loop
  end if 
  rsbody.close
  set rsbody=nothing
   
   
   
		
  

   dwt.out "</td></tr>"& vbCrLf

  
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassgz,jxgzname
   sbclassgz=conn.Execute("SELECT sb_jxgzxx_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   if not isnull( sbclassgz ) then 
	  sbclassgz1=split(sbclassgz,",")
	 For i = LBound(sbclassgz1) To UBound(sbclassgz1)
              	
				jxgzname=getjxgzxx(sbclassgz1(i))
				'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +':'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_gzxx_new' value='"&sbclassgz1(i)&"'>"	
				dwt.out i+1& "-"&jxgzname & "<br>"
				
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_gzxx_new' value='0' onclick='clickgzxxqt()'>����<span id=jxgzxxspan></span>"	
			
   dwt.out "</td></tr>"& vbCrLf
   



   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ݣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassjx,jxnrname
   sbclassjx=conn.Execute("SELECT sb_jxnr_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   
   if not isnull( sbclassjx ) then 
	  sbclassjx1=split(sbclassjx,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
                jxnrname=getjxnr(sbclassjx1(i))
				'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +':'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassjx1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_nr_new' value='"&sbclassjx1(i)&"'>"	
				dwt.out i+1& "-"&jxnrname & "<br>"
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_nr_new' value='0'  id='jxnrqt' onclick='clickjxnrqt()' >����<span id=jxnrspan></span>"	
   dwt.out"<br><input type='checkbox' name='jx_nr_new' value='9999' onclick='clickjxnrgh()'>����<span id='jxnrgh'></span>"	
   dwt.out "</td></tr>"& vbCrLf
   %>
      <script language="JavaScript">
		
		//��������ʾ�������Ƿ�ѡ��
		function clickgzxxqt(){
		  var checkName = document.getElementsByName ("jx_gzxx_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==0)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			jxgzxxspan.innerHTML="��<input name='jx_gzxx' id='jx_gzxx' type='text'>"
			}else{
			jxgzxxspan.innerHTML=""
		  }
		}
		//���������ݵ������Ƿ�ѡ��
		function clickjxnrqt(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==0)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			jxnrspan.innerHTML="��<input name='jx_nr' id='jx_nr' type='text'>"
			}else{
			jxnrspan.innerHTML=""
		  }
		}
		//���������ݵĸ����Ƿ�ѡ��
		function clickjxnrgh(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==9999)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			  <%
			  ghqxh=conn.Execute("SELECT sb_ggxh FROM sb  WHERE sb_id="&sb_id)(0)
			  
			  %>
			jxnrgh.innerHTML="������ǰ�ͺ� <input name='gh_xh' id='gh_xh'  type='text' value='<%=ghqxh%>'>&nbsp;�������ͺ�<input name='gh_xhupdate'  id='gh_xhupdate'  type='text'>"
			}else{
			jxnrgh.innerHTML=""
		  }
		}
      </script>
   <%
   
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ʱ�䣺 </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   Dwt.out"<input name='jx_date'  type='text'  id='jx_date'  class='Wdate' onFocus=""var jx_enddate=$dp.$('jx_enddate');WdatePicker({onpicked:function(){pickedFunc();jx_enddate.focus();},dateFmt:'yyyy/MM/dd HH:mm',maxDate:'#F{$dp.$D(\'jx_enddate\')}'})""   readOnly  value='"&now()&"'>"
   dwt.out " �� "
   
   Dwt.out"<input name='jx_enddate' type='text'  id='jx_enddate'  class='Wdate'   onFocus=""WdatePicker({dateFmt:'yyyy/MM/dd HH:mm',minDate:'#F{$dp.$D(\'jx_date\')}',onpicked:pickedFunc})""   readOnly  value='"&now()&"'>"
   
   dwt.out "&nbsp;&nbsp;<span id='jxys'></span>  "
   
   Dwt.out"</td></tr>"& vbCrLf
   %>
   <script language="JavaScript">
    // �����������ڵļ������  
     //   document.all.dateChangDu.value = iDays;
	function pickedFunc(){
		  Date.prototype.dateDiff = function(interval,objDate){    
		//����������� objDate �������������ش� undefined    
		if(arguments.length<2||objDate.constructor!=Date) return undefined;    
		switch (interval) {      
		//�������    
		 // case "s":return parseInt((objDate-this)/1000);      
		  //����ֲ�    
			case "n":return parseInt(Math.round(((objDate-this)/60000)*100)/100);      
			//����ʱ��    
			  case "h":return Math.round(((objDate-this)/3600000)*100)/100;      
			  //�����ղ�      
			 // case "d":return parseInt((objDate-this)/86400000);      
			  //�����²�      
			 // case "m":return (objDate.getMonth()+1)+((objDate.getFullYear()-this.getFullYear())*12)-(this.getMonth()+1);      
			  //�������      
			 // case "y":return objDate.getFullYear()-this.getFullYear();      
			
			  //��������      
			  default:return undefined;    
			}
		 }
		//document.all.dateChangDu.value = document.all.jx_date.value;
			  var sDT = new Date(document.all.jx_date.value);
			  var eDT = new Date(document.all.jx_enddate.value);
			  jxys.innerHTML=("��ʱ��"+ sDT.dateDiff("h",eDT)+"Сʱ");
	
	}

</script>  
   
   <%
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>���޸����ˣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_fzren' type='text' ></td></tr>"& vbCrLf
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�� �� �ˣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_ren' type='text'><br>�����ֵ�����,�����м�������ӿո�������ַ� <br>���������,ÿ���˵������м����ÿո�����,����ʹ�������ַ�</td></tr>"& vbCrLf
   
      Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������⣺ </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   		
					sqlz="SELECT * from sbjxylwt order by  sbjxylwt_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
						do while not rsz.eof
							dwt.out"<input type='checkbox' name='jx_ylwt' value='"&rsz("sbjxylwt_id")&"'" 
											Dwt.out ">"	
											dwt.out rsz("sbjxylwt_name") & "<br>"
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing 
  

   dwt.out "<b>��������ѡ���ʾ����������</b>"& vbCrLf
   dwt.out "<br><b>���ѡ��������Ŀ���ʾ����������,�豸""���""״̬�����Ϊ�����</b></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��    ע�� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_bz' type='text'></td></tr>"& vbCrLf
   
   
   Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
	Dwt.out"<input name='action' type='hidden' id='action' value='saveadd'> <input name='sbid' type='hidden'  value='"&Trim(Request("sbid"))&"'> <input name='sbclassid' type='hidden'  value='"&Trim(Request("sbclassid"))&"'> <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='sb_jxjl.asp?sbid="&Trim(Request("sbid"))&"&sbclassid="&Trim(Request("sbclassid"))&"';"" style='cursor:hand;'></td>  </tr>"

   Dwt.out"</form></table>"& vbCrLf
end sub	

sub saveadd()    
	'�����������޼�¼
	set rsadd=server.createobject("adodb.recordset")
	sqladd="select * from sbjx" 
	rsadd.open sqladd,conn,1,3
	rsadd.addnew
	rsadd("jx_lb")=ReplaceBadChar(Trim(Request("jx_lb")))
	rsadd("jx_gzxx")=ReplaceBadChar(Trim(Request("jx_gzxx")))
	rsadd("jx_nr")=ReplaceBadChar(Trim(request("jx_nr")))
	rsadd("jx_gzxx_new")=ReplaceBadChar(Trim(Request("jx_gzxx_new")))
	rsadd("jx_nr_new")=ReplaceBadChar(Trim(request("jx_nr_new")))
	rsadd("jx_date")=Trim(request("jx_date"))
	rsadd("jx_enddate")=Trim(request("jx_enddate"))
	rsadd("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsadd("jx_ren")=ReplaceBadChar(Trim(request("jx_ren")))
	rsadd("jx_ylwt")=ReplaceBadChar(Trim(request("jx_ylwt")))
	rsadd("jx_bz")=ReplaceBadChar(Trim(request("jx_bz")))
	rsadd("sb_id")=ReplaceBadChar(Trim(request("sbid")))
	rsadd.update
	jxid= rsadd("jx_id")
	rsadd.close
	
	
	ghxh=ReplaceBadChar(Trim(request("gh_xh")))
	ghxhupdate=ReplaceBadChar(Trim(request("gh_update")))
   dim isgh '�Ƿ�ѡ�и���
  isgh =false
   if not isnull( ReplaceBadChar(Trim(request("jx_nr_new"))) ) then 
	  sbclassjx1=split(ReplaceBadChar(Trim(request("jx_nr_new"))),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=9999 then isgh=true
	 Next 
	 
	end if  

	'�������˸�������
	if isgh and jxid<>"" then 
	  '��������������¼
      set rsaddgh=server.createobject("adodb.recordset")
      sqladdgh="select * from sbgh" 
      rsaddgh.open sqladdgh,conn,1,3
      rsaddgh.addnew
      rsaddgh("jx_id")=jxid
      rsaddgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
      rsaddgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
     rsaddgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
     rsaddgh.update
      rsaddgh.close
      set rsaddgh=nothing
	  
	   '�����豸��������Ӧλ���豸�Ĺ���ͺ�2008-9-19
	  set rsadd1=server.createobject("adodb.recordset")
          sqladd1="select * from sb where sb_id="&Trim(request("sbid"))
          rsadd1.open sqladd1,conn,1,3
          rsadd1("sb_ggxh")=ReplaceBadChar(Trim(Request("gh_xhupdate")))  
      rsadd1("sb_qydate")=ReplaceBadChar(Trim(request("gh_date")))
	  rsadd1.update
          rsadd1.close
          set rsadd1=nothing
	  
	end if 
	
      set rsadd=nothing
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sb where sb_id="&Trim(request("sbid"))
      rsedit.open sqledit,conn,1,3
      	  rsedit("sb_update")=now()
		  if ReplaceBadChar(Trim(request("jx_ylwt")))<>"" then rsedit("sb_whqk")="2"  else rsedit("sb_whqk")="1"
      rsedit.update
      rsedit.close
      set rsedit=nothing
	'rsedit("sbid")=request("sbid")
	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0))(0)
	  Dwt.savesl "�豸����-���޼�¼-"&sbclassname,"���",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0)&" ���ڣ�"&ReplaceBadChar(Trim(request("jx_date")))
	'Dwt.out"<Script Language=Javascript>history.go(-2)<Script>"
     response.write"<Script Language=Javascript>location.href='?sbid="&sb_id&"&sbclassid="&sbclass_id&"';</Script>"
end sub


sub edit()
    sqledit="SELECT * from sbjx where jx_id="&Trim(Request("jxid"))
	set rsedit=server.createobject("adodb.recordset")
    rsedit.open sqledit,conn,1,1


   Dwt.out"<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'  >"& vbCrLf
   Dwt.out"<form method='post' action='sb_jxjl.asp' name='formadd' onsubmit='javascript:return CheckAdd();'>"& vbCrLf
   Dwt.out"<tr class='title'><td height='22' colspan='2'>"& vbCrLf
   Dwt.out"<Div align='center'><strong>�༭   "&sb_wh&" ���޼�¼</strong></Div></td>    </tr>"& vbCrLf
  
  
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>������� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
    dim sqlbody,rsbody,rsz,sqlz,rszz,sqlzz
  sqlbody="SELECT * from sbjxlb where sbjxlb_zclass=0 order by  sbjxlb_orderby aSC"
  set rsbody=server.createobject("adodb.recordset")
  rsbody.open sqlbody,conn,1,1
  if rsbody.eof and rsbody.bof then 
     Dwt.out "<p align=""center"">��������</p>" 
  else
	  do while not rsbody.eof 
						'����
				sqlz="SELECT * from sbjxlb where sbjxlb_zclass="&rsbody("sbjxlb_id")&" order by  sbjxlb_orderby aSC"& vbCrLf
				set rsz=server.createobject("adodb.recordset")
				rsz.open sqlz,conn,1,1
				if rsz.eof and rsz.bof then 
					dwt.out"<input type='radio' name='jx_lb' value='"&rsbody("sbjxlb_id")&"'" 
							if not isnull(rsedit("jx_lb")) then 
							  if cint(rsedit("jx_lb"))=cint(rsbody("sbjxlb_id")) then 
								 dwt.out " checked "
							  end if    
							end if 
							Dwt.out ">"  	
							dwt.out rsbody("sbjxlb_name") & "<br>"
				else
					do while not rsz.eof
					
						dwt.out"<input type='radio' name='jx_lb' value='"&rsz("sbjxlb_id")&"'" 
							if not isnull(rsedit("jx_lb")) then 
							  if cint(rsedit("jx_lb"))=cint(rsz("sbjxlb_id")) then 
								 dwt.out " checked "
							  end if 
							end if     
								Dwt.out ">"	
								dwt.out rsbody("sbjxlb_name")&":"&rsz("sbjxlb_name") & "<br>"
					rsz.movenext
					loop
				end if 	
				rsz.close
				set rsz=nothing
			
		rsbody.movenext
		loop
  end if 
  rsbody.close
  set rsbody=nothing
   
   
   
		
  

   dwt.out "</td></tr>"& vbCrLf

  
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassgz,jxgzname
   sbclassgz=conn.Execute("SELECT sb_jxgzxx_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   if not isnull( sbclassgz ) then 
	  sbclassgz1=split(sbclassgz,",")
	 For i = LBound(sbclassgz1) To UBound(sbclassgz1)
              	jxgzname=getjxgzxx(sbclassgz1(i))

				'jxgzname=conn.Execute("SELECT sbjxgzA.sbjxgzxx_name +':'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&sbclassgz1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_gzxx_new' value='"&sbclassgz1(i)&"'"
				call checkbox(rsedit("jx_gzxx_new"),sbclassgz1(i),"")
				dwt.out">"	
				dwt.out jxgzname & "<br>"
				
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_gzxx_new' value='0' onclick='clickgzxxqt()'"
   call checkbox(rsedit("jx_gzxx_new"),0,rsedit("jx_gzxx"))
   dwt.out ">����<span id=jxgzxxspan></span>"	
			
   dwt.out "</td></tr>"& vbCrLf
   



   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������ݣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   
   
   dim sbclassjx,jxnrname
   sbclassjx=conn.Execute("SELECT sb_jxnr_class FROM sbclass WHERE sbclass_id="&sbclass_id)(0)
   
   
   
   dim is999  '�����жϵ�ǰֵ�Ƿ�����и�����¼  JS����
   is999=false
   if not isnull( rsedit("jx_nr_new") ) then 
	  sbclassjx1=split(rsedit("jx_nr_new"),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
              	if cint(sbclassjx1(i))=9999 then is999=true
   	 Next 
	end if  


   if not isnull( sbclassjx ) then 
	  sbclassjx1=split(sbclassjx,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
                jxnrname=getjxnr(sbclassjx1(i))
				'jxnrname=conn.Execute("SELECT sbjxnrA.sbjxnr_name +':'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&sbclassjx1(i))(0)
			    dwt.out"<input type='checkbox' name='jx_nr_new' value='"&sbclassjx1(i)&"' "
				call checkbox(rsedit("jx_nr_new"),sbclassjx1(i),"")

				dwt.out " >"	
				dwt.out jxnrname & "<br>"
   	 Next 
	end if  
   dwt.out"<input type='checkbox' name='jx_nr_new' value='0'  id='jxnrqt' onclick='clickjxnrqt()' "
	call checkbox(rsedit("jx_nr_new"),0,rsedit("jx_nr"))
   dwt.out " >����<span id=jxnrspan></span>"	
   dwt.out"<br><input type='checkbox' name='jx_nr_new' value='9999' onclick='clickjxnrgh()'"
	call checkbox(rsedit("jx_nr_new"),9999,"")
   dwt.out">����<span id='jxnrgh'></span>"	

   dwt.out "</td></tr>"& vbCrLf
   %>
      <script language="JavaScript">
		
		//��������ʾ�������Ƿ�ѡ��
		function clickgzxxqt(){
		  var checkName = document.getElementsByName ("jx_gzxx_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==0)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			jxgzxxspan.innerHTML="��<input name='jx_gzxx' id='jx_gzxx' type='text'  value='<%=rsedit("jx_gzxx")%>'>"
			}else{
			jxgzxxspan.innerHTML=""
		  }
		}
		//���������ݵ������Ƿ�ѡ��
		function clickjxnrqt(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==0)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			jxnrspan.innerHTML="��<input name='jx_nr' id='jx_nr' type='text'  value='<%=rsedit("jx_nr")%>'>"
			}else{
			jxnrspan.innerHTML=""
		  }
		}
		//���������ݵĸ����Ƿ�ѡ��
		function clickjxnrgh(){
		  var checkName = document.getElementsByName ("jx_nr_new");	//�����������ȡ�齨����
		  var ischecked
		  //ѭ��checkbox���ж��Ƿ����ѡ����
		  for (i = 0; i < checkName.length; i ++) {
			  if (checkName[i].checked) {	//�����ѡ����򷵻�true
				  
				  if(checkName[i].value==9999)ischecked= true;   //���ѡ�е���0,Ҳ����"����"�����
			  }
		  }   
		  if(ischecked){
			  <%
			  ghqxh=conn.Execute("SELECT sb_ggxh FROM sb  WHERE sb_id="&sb_id)(0)
			  
			  %>
			  <%
			  if is999 then 
			     ghqxh=conn.Execute("SELECT gh_xh FROM sbgh  WHERE jx_id="&Trim(Request("jxid")))(0)
			    'dwt.out "sfsdfdsfd"
			     ghxhupdate=conn.Execute("SELECT gh_xhupdate FROM sbgh  WHERE jx_id="&Trim(Request("jxid")))(0)
			  end if 
			  %>
			  

			jxnrgh.innerHTML="������ǰ�ͺ� <input name='gh_xh' id='gh_xh'  type='text' value='<%=ghqxh%>'>&nbsp;�������ͺ�<input name='gh_xhupdate'  id='gh_xhupdate'  value='<%=ghxhupdate%>' type='text'>"
			}else{
			jxnrgh.innerHTML=""
		  }
		}
      </script>
   <%
   
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>����ʱ�䣺 </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   Dwt.out"<input name='jx_date'  type='text'  id='jx_date'  class='Wdate' onFocus=""var jx_enddate=$dp.$('jx_enddate');WdatePicker({onpicked:function(){pickedFunc();jx_enddate.focus();},dateFmt:'yyyy/MM/dd HH:mm',maxDate:'#F{$dp.$D(\'jx_enddate\')}'})""   readOnly  value='"&rsedit("jx_date")&"'>"
   dwt.out " �� "
   
   Dwt.out"<input name='jx_enddate' type='text'  id='jx_enddate'  class='Wdate'   onFocus=""WdatePicker({dateFmt:'yyyy/MM/dd HH:mm',minDate:'#F{$dp.$D(\'jx_date\')}',onpicked:pickedFunc})""   readOnly  value='"&rsedit("jx_enddate")&"'>"
   
   dwt.out "&nbsp;&nbsp;<span id='jxys'></span>  "
   
   Dwt.out"</td></tr>"& vbCrLf
   %>
   <script language="JavaScript">
    // �����������ڵļ������  
     //   document.all.dateChangDu.value = iDays;
	function pickedFunc(){
		  Date.prototype.dateDiff = function(interval,objDate){    
		//����������� objDate �������������ش� undefined    
		if(arguments.length<2||objDate.constructor!=Date) return undefined;    
		switch (interval) {      
		//�������    
		 // case "s":return parseInt((objDate-this)/1000);      
		  //����ֲ�    
			case "n":return parseInt(Math.round(((objDate-this)/60000)*100)/100);      
			//����ʱ��    
			  case "h":return Math.round(((objDate-this)/3600000)*100)/100;      
			  //�����ղ�      
			 // case "d":return parseInt((objDate-this)/86400000);      
			  //�����²�      
			 // case "m":return (objDate.getMonth()+1)+((objDate.getFullYear()-this.getFullYear())*12)-(this.getMonth()+1);      
			  //�������      
			 // case "y":return objDate.getFullYear()-this.getFullYear();      
			
			  //��������      
			  default:return undefined;    
			}
		 }
		//document.all.dateChangDu.value = document.all.jx_date.value;
			  var sDT = new Date(document.all.jx_date.value);
			  var eDT = new Date(document.all.jx_enddate.value);
			  jxys.innerHTML=("��ʱ��"+ sDT.dateDiff("h",eDT)+"Сʱ");
	
	}

</script>  
   
   <%
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>���޸����ˣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_fzren' type='text'  value='"&rsedit("jx_fzren")&"' ></td></tr>"& vbCrLf
   
   Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�� �� �ˣ� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_ren' type='text' value='"&rsedit("jx_ren")&"'><br>�����ֵ�����,�����м�������ӿո�������ַ� <br>���������,ÿ���˵������м����ÿո�����,����ʹ�������ַ�</td></tr>"& vbCrLf
   
      Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>�������⣺ </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'>"
   		
					sqlz="SELECT * from sbjxylwt order by  sbjxylwt_orderby aSC"
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
						do while not rsz.eof
							dwt.out"<input type='checkbox' name='jx_ylwt' value='"&rsz("sbjxylwt_id")&"'" 
							dwt.out rsedit("jx_ylwt")&"-"&rsz("sbjxylwt_id")
					       
						   dwt.out checkbox(rsedit("jx_ylwt"),rsz("sbjxylwt_id"),"")
								
							Dwt.out ">"	
							dwt.out rsz("sbjxylwt_name") & "<br>"
						rsz.movenext
						loop
					end if 	
					rsz.close
					set rsz=nothing 
  

   dwt.out "<b>��������ѡ���ʾ����������</b>"& vbCrLf
   dwt.out "<br><b>���ѡ��������Ŀ���ʾ����������,�豸""���""״̬�����Ϊ�����</b></td></tr>"& vbCrLf
     Dwt.out"<tr class='tdbg'><td width='20%' align='right' class='tdbg'><strong>��    ע�� </strong></td>"   & vbCrLf   
   Dwt.out"<td width='80%' class='tdbg'><input name='jx_bz' type='text' value='"&rsedit("jx_bz")&"'></td></tr>"& vbCrLf
   
   
   Dwt.out"<tr><td height='40' colspan='2' align='center' class='tdbg'>"
   Dwt.out"<input name='action' type='hidden' id='action' value='saveedit'>   <input name='sbid' type='hidden'  value='"&Trim(Request("sbid"))&"'> <input name='sbclassid' type='hidden'  value='"&Trim(Request("sbclassid"))&"'>  <input type='hidden' name='jxid' value='"&Trim(Request("jxid"))&"'>     <input  type='submit' name='Submit' value=' �� �� ' style='cursor:hand;'>&nbsp;<input name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick=""window.location.href='sb_jxjl.asp?sbid="&Trim(Request("sbid"))&"&sbclassid="&Trim(Request("sbclassid"))&"';"" style='cursor:hand;'></td>  </tr>"

   Dwt.out"</form></table>"& vbCrLf























   rsedit.close
   set rsedit=nothing
end sub

'�������ƣ�checkbox ҳ���Ƿ�ѡ��
'���ã��жϼ��޼�¼�����ݺ��豸����ļ��������Ƿ��Ӧ ���޼�¼�к��豸�����еļ������ݶ�Ӧ�������checked

'jx_gzxx_new  ����,���޼�¼�б����ֵ
'sbclassjxid  ���õ�ʱ���Ѿ��ָ��  �豸�����еļ��޼�¼����
'jx_gzxx  ���ݾɵ�����,����д���Ϣ,��"����"��Ĭ��ѡ�е�
Function checkbox(jx_gzxx_new,sbclassjxid,jx_gzxx)
	dim sbclassjx1,i
	if not isnull( jx_gzxx_new ) then 
	  sbclassjx1=split(jx_gzxx_new,",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=cint(sbclassjxid) then dwt.out " checked "
	 Next 
	 
	end if  
	 if jx_gzxx<>"" then dwt.out " checked "
end Function


sub saveedit()
	jxid=ReplaceBadChar(Trim(request("jxID")))
	
	'�༭����
	set rsedit=server.createobject("adodb.recordset")
	sqledit="select * from sbjx where jx_ID="&jxid
	
	rsedit.open sqledit,conn,1,3







	rsedit("jx_lb")=ReplaceBadChar(Trim(Request("jx_lb")))
	rsedit("jx_gzxx")=ReplaceBadChar(Trim(Request("jx_gzxx")))
	rsedit("jx_nr")=ReplaceBadChar(Trim(request("jx_nr")))
	rsedit("jx_gzxx_new")=ReplaceBadChar(Trim(Request("jx_gzxx_new")))
	rsedit("jx_nr_new")=ReplaceBadChar(Trim(request("jx_nr_new")))
	rsedit("jx_date")=Trim(request("jx_date"))
	rsedit("jx_enddate")=Trim(request("jx_enddate"))
	rsedit("jx_fzren")=ReplaceBadChar(Trim(request("jx_fzren")))
	rsedit("jx_ren")=ReplaceBadChar(Trim(request("jx_ren")))
	rsedit("jx_ylwt")=ReplaceBadChar(Trim(request("jx_ylwt")))
	rsedit("jx_bz")=ReplaceBadChar(Trim(request("jx_bz")))
	rsedit("sb_id")=ReplaceBadChar(Trim(request("sbid")))
	
      rsedit.update
      rsedit.close
      set rsedit=nothing
	
	ghxh=ReplaceBadChar(Trim(request("gh_xh")))
	ghxhupdate=ReplaceBadChar(Trim(request("gh_update")))


   dim isgh '�Ƿ�ѡ�и���
  isgh =false
   if not isnull( ReplaceBadChar(Trim(request("jx_nr_new"))) ) then 
	  sbclassjx1=split(ReplaceBadChar(Trim(request("jx_nr_new"))),",")
	 For i = LBound(sbclassjx1) To UBound(sbclassjx1)
		if cint(sbclassjx1(i))=9999 then isgh=true
	 Next 
	 
	end if  

	'���ѡ���˸�������,�򱣴����Ϣ,���ûѡ�� ���жϸ�����¼����û��ЩJXID������,�еĻ�ɾ��
	if isgh and jxid<>"" then 
 
                    '����Ƿ���ЩJXID�ĸ�����¼,û�о�����,�о͸���
					sqlz="SELECT * from sbgh where jx_id="&jxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
						'��������������¼
						set rsaddgh=server.createobject("adodb.recordset")
						sqladdgh="select * from sbgh" 
						rsaddgh.open sqladdgh,conn,1,3
						rsaddgh.addnew
						rsaddgh("jx_id")=jxid
						rsaddgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
						rsaddgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
					   rsaddgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
					   rsaddgh.update
						rsaddgh.close
						set rsaddgh=nothing
					else
					  set rseditgh=server.createobject("adodb.recordset")
					  sqleditgh="select * from sbgh where jx_ID="&jxid
					  
					  rseditgh.open sqleditgh,conn,1,3
				  
						rseditgh("gh_xh")=ReplaceBadChar(Trim(Request("gh_xh")))
						rseditgh("gh_xhupdate")=ReplaceBadChar(Trim(Request("gh_xhupdate")))
					   'rseditgh("sb_id")=ReplaceBadChar(Trim(request("sbid")))
				  
						rseditgh.update
						rseditgh.close
						set rsedight=nothing

					end if 	
					rsz.close
					set rsz=nothing 


	 
	 
	  
			 '�����豸��������Ӧλ���豸�Ĺ���ͺ�2008-9-19
			set rsadd1=server.createobject("adodb.recordset")
				sqladd1="select * from sb where sb_id="&Trim(request("sbid"))
				rsadd1.open sqladd1,conn,1,3
				rsadd1("sb_ggxh")=ReplaceBadChar(Trim(Request("gh_xhupdate")))  
			rsadd1("sb_qydate")=ReplaceBadChar(Trim(request("gh_date")))
			rsadd1.update
				rsadd1.close
				set rsadd1=nothing
	  
	else
'ɾ������ ��Ϣ
					sqlz="SELECT * from sbgh where jx_id="&jxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					else
					  set rsdel=server.createobject("adodb.recordset")
					  sqldel="delete * from sbgh where jx_id="&jxid
					  rsdel.open sqldel,conn,1,3
					  'rsdel.close
					  set rsdel=nothing  
					 
					 

					end if 	
					rsz.close
					set rsz=nothing 
	end if 
	
	   
      set rsedit=server.createobject("adodb.recordset")
      sqledit="select * from sb where sb_id="&Trim(request("sbid"))
      rsedit.open sqledit,conn,1,3
      	  rsedit("sb_update")=now()
		  if ReplaceBadChar(Trim(request("jx_ylwt")))<>"" then rsedit("sb_whqk")="2"  else rsedit("sb_whqk")="1"
      rsedit.update
      rsedit.close
      set rsedit=nothing



















	

	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0))(0)
	  Dwt.savesl "�豸����-���޼�¼-"&sbclassname,"�༭",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&ReplaceBadChar(Trim(request("sbid"))))(0)&" ���ڣ�"&ReplaceBadChar(Trim(request("jx_date")))

	Dwt.out"<Script Language=Javascript>history.go(-2)</Script>"
end sub


sub del()
	jx_ID=request("jxID")
	sb_id=conn.Execute("SELECT sb_id FROM sbjx WHERE jx_id="&jx_id)(0)
	  sbclassname=conn.Execute("SELECT sbclass_name FROM sbclass WHERE sbclass_id="&conn.Execute("SELECT sb_dclass FROM sb WHERE sb_id="&sb_id)(0))(0)
	deljxdate=conn.Execute("SELECT jx_date FROM sbjx WHERE jx_id="&jx_id)(0)
	  Dwt.savesl "�豸����-���޼�¼-"&sbclassname,"ɾ��",conn.Execute("SELECT sb_wh FROM sb WHERE sb_id="&sb_id)(0)&" ʱ�䣺"&deljxdate
	
	set rsdel=server.createobject("adodb.recordset")
	sqldel="delete * from sbjx where jx_id="&jx_id
	rsdel.open sqldel,conn,1,3
	Dwt.out"<Script Language=Javascript>history.back()</Script>"
	'rsdel.close
	set rsdel=nothing  

end sub


'**********************************************
'��¼���������Ƿ����޸�Ȩ��ҳ��jxjl.asp
'******************************8
sub jxxuanxiang(id,sb_id,sb_sscj)
 if session("levelclass")=sb_sscj or session("levelclass")=0 then 
	Dwt.out"<a href=sb_jxjl.asp?action=edit&sbid="&sb_id&"&sbclassid="&sbclass_id&"&jxid="&rsjx("jx_id")&">��</a>&nbsp;"
	Dwt.out"<a href=sb_jxjl.asp?action=del&jxid="&rsjx("jx_id")&"&sbclassid="&sbclass_id&"&sbid="&sb_id&" onClick=""return confirm('ȷ��Ҫɾ���˼�¼��');"">ɾ</a>"
 else
    Dwt.out"&nbsp;"
 end if 
end sub



Call CloseConn
%>