<!--#include file="conn.asp"-->
<%

'��ȡ �������ݵ�����
function getjxnr(jxnrid)
				'������ ͷsqlz="SELECT sbjxnrA.sbjxnr_name +'��'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&jxnrid
				sqlz="SELECT  sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&jxnrid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxnr="���ֶ����ݱ�ɾ��"
					else
						getjxnr=rsz("sbjxnr_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'��ȡ ���޹�����Ϣ������
function getjxgzxx(jxgzxxid)
				'����ͷsqlz="SELECT sbjxgzA.sbjxgzxx_name +'��'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&jxgzxxid
				sqlz="SELECT sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&jxgzxxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxgzxx="���ֶ����ݱ�ɾ��"
					else
						getjxgzxx=rsz("sbjxgzxx_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'��ȡ ������������
function getjxlb(jxlbid)
				'���� ͷsqlz="SELECT iif(sbjxlb.sbjxlb_zclass<>0,sbjxlbA.sbjxlb_name+'��'+sbjxlb.sbjxlb_name,sbjxlb.sbjxlb_name) as sbjxlb_name  FROM sbjxlb AS sbjxlb left join sbjxlb as sbjxlbA on sbjxlb.sbjxlb_zclass=sbjxlbA.sbjxlb_id WHERE sbjxlb.sbjxlb_id="&jxlbid
				sqlz="SELECT iif(sbjxlb.sbjxlb_zclass<>0,sbjxlb.sbjxlb_name,sbjxlb.sbjxlb_name) as sbjxlb_name  FROM sbjxlb AS sbjxlb left join sbjxlb as sbjxlbA on sbjxlb.sbjxlb_zclass=sbjxlbA.sbjxlb_id WHERE sbjxlb.sbjxlb_id="&jxlbid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxlb="���ֶ����ݱ�ɾ��"
					else
						getjxlb=rsz("sbjxlb_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'��ȡ �����������������
function getjxylwt(jxylwtid)
				sqlz="SELECT  sbjxylwt.sbjxylwt_name as sbjxylwt_name FROM sbjxylwt AS sbjxylwt left join sbjxylwt as sbjxylwtA on sbjxylwt.sbjxylwt_zclass=sbjxylwtA.sbjxylwt_id WHERE sbjxylwt.sbjxylwt_id="&jxylwtid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxylwt="���ֶ����ݱ�ɾ��"
					else
						getjxylwt=rsz("sbjxylwt_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

%>