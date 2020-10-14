<%Function toUTF8(szInput)'GB转UTF8--将GB编码文字转换为UTF8编码文字

    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    '如果输入参数为空，则退出函数
    If szInput = "" Then
        toUTF8 = szInput
        Exit Function
    End If
    '开始转换
     For x = 1 To Len(szInput)
        '利用mid函数分拆GB编码文字
        wch = Mid(szInput, x, 1)
        '利用ascW函数返回每一个GB编码文字的Unicode字符代码
        '注：asc函数返回的是ANSI 字符代码，注意区别
        nAsc = AscW(wch)
        If nAsc < 0 Then nAsc = nAsc + 65536
    
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
               'GB编码文字的Unicode字符代码在0800 - FFFF之间采用三字节模版
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                            Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                            Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
        
    toUTF8 = szRet
End Function





function UTF2GB(UTFStr) 'UTF转GB---将UTF8编码文字转换为GB编码文字


for Dig=1 to len(UTFStr) 
  '如果UTF8编码文字以%开头则进行转换
  if mid(UTFStr,Dig,1)="%" then 
     'UTF8编码文字大于8则转换为汉字
    if len(UTFStr) >= Dig+8 then 
       GBStr=GBStr & ConvChinese(mid(UTFStr,Dig,9)) 
       Dig=Dig+8 
    else 
      GBStr=GBStr & mid(UTFStr,Dig,1) 
    end if 
  else 
     GBStr=GBStr & mid(UTFStr,Dig,1) 
  end if 
next 
UTF2GB=GBStr 
end function 

'UTF8编码文字将转换为汉字
function ConvChinese(x) 
   A=split(mid(x,2),"%") 
   i=0 
   j=0 
  for i=0 to ubound(A) 
     A(i)=c16to2(A(i)) 
  next 
  for i=0 to ubound(A)-1 
    DigS=instr(A(i),"0") 
    Unicode="" 
    for j=1 to DigS-1 
      if j=1 then 
        A(i)=right(A(i),len(A(i))-DigS) 
        Unicode=Unicode & A(i) 
      else 
         i=i+1 
         A(i)=right(A(i),len(A(i))-2) 
         Unicode=Unicode & A(i) 
      end if 
    next 

    if len(c2to16(Unicode))=4 then 
       ConvChinese=ConvChinese & chrw(int("&H" & c2to16(Unicode))) 
    else 
       ConvChinese=ConvChinese & chr(int("&H" & c2to16(Unicode))) 
    end if 
  next 
end function 

'二进制代码转换为十六进制代码
function c2to16(x)
   i=1 
   for i=1 to len(x) step 4 
      c2to16=c2to16 & hex(c2to10(mid(x,i,4))) 
   next 
end function 

'二进制代码转换为十进制代码
function c2to10(x)
   c2to10=0 
   if x="0" then exit function 
     i=0 
   for i= 0 to len(x) -1 
      if mid(x,len(x)-i,1)="1" then c2to10=c2to10+2^(i) 
   next 
end function 

'十六进制代码转换为二进制代码
function c16to2(x) 
    i=0 
    for i=1 to len(trim(x)) 
      tempstr= c10to2(cint(int("&h" & mid(x,i,1)))) 
      do while len(tempstr)<4 
         tempstr="0" & tempstr 
      loop 
      c16to2=c16to2 & tempstr 
   next 
end function 

'十进制代码转换为二进制代码
function c10to2(x) 
   mysign=sgn(x) 
   x=abs(x) 
   DigS=1 
   do 
      if x<2^DigS then 
        exit do 
      else 
        DigS=DigS+1 
      end if 
   loop 
   tempnum=x 

   i=0 
   for i=DigS to 1 step-1 
      if tempnum>=2^(i-1) then 
         tempnum=tempnum-2^(i-1) 
         c10to2=c10to2 & "1" 
      else 
         c10to2=c10to2 & "0" 
      end if 
   next 
   if mysign=-1 then c10to2="-" & c10to2 
end function
%>