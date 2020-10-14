<%
Function getHTTPPage(url)
    On Error Resume Next
    Dim http
    Set http = Server.CreateObject("Microsoft.XMLHTTP")
    Http.Open "GET", url, false
    Http.send()
    If Http.readystate<>4 Then
        Exit Function
    End If
    getHTTPPage = bytes2BSTR(Http.responseBody)
    Set http = Nothing
    If Err.Number<>0 Then Err.Clear
End Function

Function bytes2BSTR(vIn)
    Dim strReturn
    Dim i1, ThisCharCode, NextCharCode
    strReturn = ""
    For i1 = 1 To LenB(vIn)
        ThisCharCode = AscB(MidB(vIn, i1, 1))
        If ThisCharCode < &H80 Then
            strReturn = strReturn & Chr(ThisCharCode)
        Else
            NextCharCode = AscB(MidB(vIn, i1 + 1, 1))
            strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
            i1 = i1 + 1
        End If
    Next
    bytes2BSTR = strReturn
End Function

Function URLDecode(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = Eval("&h" + Mid(enStr, i + 1, 2))
            If v<128 Then
                deStr = deStr&Chr(v)
                i = i + 2
            Else
                If isvalidhex(Mid(enstr, i, 3)) Then
                    If isvalidhex(Mid(enstr, i + 3, 3)) Then
                        v = Eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr = deStr&Chr(v)
                        i = i + 5
                    Else
                        v = Eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr = deStr&Chr(v)
                        i = i + 3
                    End If
                Else
                    destr = destr&c
                End If
            End If
        Else
            If c = "+" Then
                deStr = deStr&" "
            Else
                deStr = deStr&c
            End If
        End If
    Next
    URLDecode = deStr
End Function

Function isvalidhex(Str)
    isvalidhex = true
    Str = UCase(Str)
    If Len(Str)<>3 Then isvalidhex = false
    Exit Function
    If Left(Str, 1)<>"%" Then isvalidhex = false
    Exit Function
    c = Mid(Str, 2, 1)
    If Not (((c>= "0") And (c<= "9")) Or ((c>= "A") And (c<= "Z"))) Then isvalidhex = false
    Exit Function
    c = Mid(Str, 3, 1)
    If Not (((c>= "0") And (c<= "9")) Or ((c>= "A") And (c<= "Z"))) Then isvalidhex = false
    Exit Function
End Function

Dim Weather, City, Start_Str, End_Str, Return_Str


 City = "长治"
 City1 = "太原"
 
Weather = getHTTPPage("http://weather.tq121.com.cn/mapanel/index_new.php?city="&City)
'weather1=getHTTPPage("http://weather.tq121.com.cn/mapanel/index_new.php?city="&City1)
If InStr(Weather, "未找到您查的城市") Then
    Return_Str = "未找到您查的城市"
Else
    Start_Str = InStr(Weather, "<hr width=""100%"" size=""1"">") + 60
    End_Str = InStr(Start_Str, Weather, "<hr width=""100%"" size=""1"">") -145
    Return_Str = Mid(Weather, Start_Str, End_Str - Start_Str)
End If



'If InStr(Weather1, "未找到您查的城市") Then
'    Return_Str1 = "未找到您查的城市"
'Else
'    Start_Str1 = InStr(Weather1, "<hr width=""100%"" size=""1"">") + 60
''    End_Str1 = InStr(Start_Str1, Weather1, "<hr width=""100%"" size=""1"">") -145
'    Return_Str1 = Mid(Weather, Start_Str1, End_Str1 - Start_Str1)
'End If


Dim Re
Set Re = New RegExp
Re.IgnoreCase = true
Re.Global = True
Re.Pattern = "\s"
Return_Str = Re.Replace(Return_Str, "")
Re.Pattern = "<tablewidth=""166""height=""15""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""160""align=""center""valign=""top""class=""weather"">([^>]+)</td></tr></table>"
FilterStr = Re.Replace(Return_Str, " $1 ")
Re.Pattern = "<tablewidth=""166""height=""28""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""160""align=""center""valign=""top""class=""weatheren"">([^>]+)</td></tr></table>"
FilterStr = Re.Replace(FilterStr, " $1 ")
Re.Pattern = "<tablewidth=""169""height=""63""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""16"">&nbsp;</td><tdwidth=""153""valign=""top""><spanclass=""big-cn"">(.+)</span></td></tr></table>"
FilterStr = Re.Replace(FilterStr, " $1 ")
FilterStr = Replace(FilterStr, "<br>", "  ")

'Dim Re1
'Set Re1 = New RegExp
'Re1.IgnoreCase = true
'Re1.Global = True
'Re1.Pattern = "\s"

'Return_Str1 = Re1.Replace(Return_Str1, "")
'Re1.Pattern = "<tablewidth=""166""height=""15""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""160""align=""center""valign=""top""class=""weather"">([^>]+)</td></tr></table>"
'FilterStr1 = Re1.Replace(Return_Str1, " $1 ")
'Re1.Pattern = "<tablewidth=""166""height=""28""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""160""align=""center""valign=""top""class=""weatheren"">([^>]+)</td></tr></table>"
'FilterStr1 = Re1.Replace(FilterStr1, " $1 ")
'Re1.Pattern = "<tablewidth=""169""height=""63""border=""0""cellpadding=""0""cellspacing=""0""><tr><tdwidth=""16"">&nbsp;</td><tdwidth=""153""valign=""top""><spanclass=""big-cn"">(.+)</span></td></tr></table>"
'FilterStr1 = Re1.Replace(FilterStr1, " $1 ")
'FilterStr1 = Replace(FilterStr1, "<br>", "  ")


Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
Response.Charset = "gb2312"
If Return_Str = "未找到您查的城市" Then
    Response.Write "返回错误："&FilterStr
Else
    Response.Write City&"今日："&FilterStr
	'Response.Write City1&"今日天气："&FilterStr1
End If






%>
