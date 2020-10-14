<!--#include file="conn.asp"-->
<%

'获取 检修内容的内容
function getjxnr(jxnrid)
				'带有题 头sqlz="SELECT sbjxnrA.sbjxnr_name +'：'+ sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&jxnrid
				sqlz="SELECT  sbjxnr.sbjxnr_name as sbjxnr_name FROM sbjxnr AS sbjxnr left join sbjxnr as sbjxnrA on sbjxnr.sbjxnr_zclass=sbjxnrA.sbjxnr_id WHERE sbjxnr.sbjxnr_id="&jxnrid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxnr="此字段内容被删除"
					else
						getjxnr=rsz("sbjxnr_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'获取 检修故障信息的内容
function getjxgzxx(jxgzxxid)
				'有题头sqlz="SELECT sbjxgzA.sbjxgzxx_name +'：'+ sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&jxgzxxid
				sqlz="SELECT sbjxgz.sbjxgzxx_name as sbjxgzxx_name FROM sbjxgzxx AS sbjxgz left join sbjxgzxx as sbjxgzA on sbjxgz.sbjxgzxx_zclass=sbjxgzA.sbjxgzxx_id WHERE sbjxgz.sbjxgzxx_id="&jxgzxxid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxgzxx="此字段内容被删除"
					else
						getjxgzxx=rsz("sbjxgzxx_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'获取 检修类别的内容
function getjxlb(jxlbid)
				'有题 头sqlz="SELECT iif(sbjxlb.sbjxlb_zclass<>0,sbjxlbA.sbjxlb_name+'：'+sbjxlb.sbjxlb_name,sbjxlb.sbjxlb_name) as sbjxlb_name  FROM sbjxlb AS sbjxlb left join sbjxlb as sbjxlbA on sbjxlb.sbjxlb_zclass=sbjxlbA.sbjxlb_id WHERE sbjxlb.sbjxlb_id="&jxlbid
				sqlz="SELECT iif(sbjxlb.sbjxlb_zclass<>0,sbjxlb.sbjxlb_name,sbjxlb.sbjxlb_name) as sbjxlb_name  FROM sbjxlb AS sbjxlb left join sbjxlb as sbjxlbA on sbjxlb.sbjxlb_zclass=sbjxlbA.sbjxlb_id WHERE sbjxlb.sbjxlb_id="&jxlbid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxlb="此字段内容被删除"
					else
						getjxlb=rsz("sbjxlb_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

'获取 检修遗留问题的内容
function getjxylwt(jxylwtid)
				sqlz="SELECT  sbjxylwt.sbjxylwt_name as sbjxylwt_name FROM sbjxylwt AS sbjxylwt left join sbjxylwt as sbjxylwtA on sbjxylwt.sbjxylwt_zclass=sbjxylwtA.sbjxylwt_id WHERE sbjxylwt.sbjxylwt_id="&jxylwtid
					set rsz=server.createobject("adodb.recordset")
					rsz.open sqlz,conn,1,1
					if rsz.eof and rsz.bof then 
					  getjxylwt="此字段内容被删除"
					else
						getjxylwt=rsz("sbjxylwt_name")
					end if 	
					rsz.close
					set rsz=nothing 
end function

%>