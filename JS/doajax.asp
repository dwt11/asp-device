<%
Dim RetCode,RetDes

function CreateXML()
  Dim OutStr
  OutStr="<?xml version=""1.0"" encoding=""gb2312""?>"&vbcrlf
  OutStr=OutStr&"<ReturnStr>"&vbcrlf
  OutStr=OutStr&"<RetCode>"&RetCode&"</RetCode>"&vbcrlf
  OutStr=OutStr&"<RetDes>"&RetDes&"</RetDes>"&vbcrlf
  OutStr=OutStr&"</ReturnStr>"
  Response.ContentType="text/xml"
  Response.write OutStr
end function


if request("action")="" then
	RetCode="0001"
	RetDes="异常错误"
	CreateXML()
response.end
end if
select case Lcase(trim(request("action")))
	case "disply"
	dim sd_dclassid
	sd_dclassid=request("sd_dclassid")
'		if Session("_WUserID") = "" then	'判断是否登入
'			RetCode="0002"
'			Conn.Execute("Update WoWo_Source Set Src_HitNum=Src_HitNum+1,Src_HitUpdate='"&Now()&"' Where Src_ID="&src_id)
'			RetDes="顶成功,但您未登陆信息无法长久保存"
'		else
'			Is_Hit_Temp =Conn.Execute("Select Count(Hit_ID) From WoWo_SrcHit Where Hit_SrcID="&Src_ID&" and Hit_UserID="&Session("_WUserID"))(0)
'			if Is_Hit_Temp <= 0 then	'判断是否顶完(避免开多个窗口的问题)	
'				Sql_Hit = "Insert into WoWo_SrcHit(Hit_SrcID,Hit_UserID,Hit_Time,Hit_IP)"
'				Sql_Hit = Sql_Hit & "Values(" & src_id & ",'" & Session("_WUserID") & "','" & Now() & "','" & Request.ServerVariables("REMOTE_ADDR") & "')"
'				Conn.Execute(Sql_Hit)
'				Conn.Execute("Update WoWo_Source Set Src_HitNum=Src_HitNum+1,Src_HitUpdate='"&Now()&"' Where Src_ID="&src_id)
				RetCode="0000"
				RetDes="顶成功,谢谢参于"
'			else
'				RetCode="0003"
'				RetDes="已顶"
'			end if
'		end if
  case else
  	RetCode="0001"
  	RetDes="异常错误"
end select
conn.close
set conn=nothing
CreateXML()
%>