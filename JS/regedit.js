
////usermanagement.asp ע���û�ͨ�Ű���֤
var num=3;
var mon1=0;
var mon2=0;
var mon3=0;
var mon4=0;
var mon5=0;

//��װ�õ�����ID����
function getObj(objName){return(document.getElementById(objName));}

//input�õ�����Ч��
function showare(id){
        for(var i=1;i<=num;i++){
		if(i==id) {
		 getObj("sps"+i).className="reg2";
	}
		else{

		  if(i==num || i==(id-1)){
		     getObj("sps"+i).className="reg1";
			 }
		  else{
			 getObj("sps"+i).className="reg1";
			 }
		 }
	}
}

//����
function RTrim(str)
{
 var whitespace=new String(" \t\n\r");
 var s=new String(str);
 if (whitespace.indexOf(s.charAt(s.length-1))!=-1)
 {
  var i=s.length-1;
  while (i>=0 && whitespace.indexOf(s.charAt(i))!=-1)
  {
   i--;
  }
    s=s.substring(0,i+1);
  }
 return s;
}

//����
function LTrim(str)
{
 var whitespace=new String(" \t\n\r");
 var s=new String(str);
 if (whitespace.indexOf(s.charAt(0))!=-1)
 {
  var j=0, i = s.length;
  while (j<i && whitespace.indexOf(s.charAt(j))!=-1)
 {
  j++;
 }
   s=s.substring(j,i);
 }
  return s;
}

function Trim(str)
{
 return RTrim(LTrim(str));
}

//����������
function checkpassword()
{ 

 if (getObj("input2").value=="" || Trim(getObj("input2").value)=="")
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>������ע�����룡</font>";

   mon2=0;
   allok();
   return false;
 }else{
 if (Trim(getObj("input2").value).indexOf(" ")>=0) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>�����в��ܰ����ո�</font>";

   mon2=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input2").value).length<4) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>���벻������4���ַ���</font>";

   mon2=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input2").value).length>20) 
 { 
   getObj("sps2").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>���벻�ܳ���20���ַ���</font>";

   mon2=0;
   allok();
  return false;
 }else{
   getObj("sps2").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>��������OK����</font>"; 
   mon2=1;
   allok(); 
  return false;
}
 }
 }
 }
}

//�ظ�����������
function checkreturnpass()
{ 

 if (getObj("input3").value=="" || Trim(getObj("input3").value)=="")
 { 
   getObj("sps3").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>��ȷ�����룡</font>";

   mon3=0;
   allok();
   return false;
 }else{
   if(getObj("input2").value!=getObj("input3").value)
   {
     getObj("sps3").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>��������������벻�����</font>"

   mon3=0;
   allok();
     return false;
   }else{
    getObj("sps3").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>����������ȷ��</font>"
   mon3=1;
   allok();   
  return false; 
}
}
}

//Ȩ�޼��
function checklevelclass(){
  if (getObj("input5").value=="" || Trim(getObj("input5").value)=="")
 { 
   getObj("sps5").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>��ѡ��Ȩ�ޣ�</font>";

   mon5=0;
   allok();
   return false;
 }else{
    getObj("sps5").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>Ȩ����ѡ��</font>"
   mon5=1;
   allok();   
  return false; 
}
}

////���ּ��
//function checkusername1(){
//  if (getObj("input4").value=="" || Trim(getObj("input4").value)=="")
// { 
//   getObj("sps4").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>����д��</font>";
//
//   mon4=0;
//   allok();
//   return false;
// }else{
//    getObj("sps4").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>Ȩ����ѡ��</font>"
//   mon4=1;
//   allok();   
//  return false; 
//}
//}


//�û������
function myuser()
{
 if (getObj("input1").value=="" || Trim(getObj("input1").value)=="")
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>�������û�����</font>";

   mon1=0;
   allok();
   return false;
 }else{
 if (Trim(getObj("input1").value).indexOf(" ")>=0) 
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>�û����в��ܰ����ո�</font>";

   mon1=0;
   allok();
  return false; 
 }else{
 if (Trim(getObj("input1").value).length<3) 
 { 
   getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>�û�����������3���ַ���</font>";

   mon1=0;
   allok();
  return false; 
 }else{	var username=getObj("input1").value;
	 //alert username;
	/*
	*--------------- GetResult() -----------------
	* GetResult()
	* ����:ͨ��XMLHTTP��������,���ؽ��.
	*--------------- GetResult() -----------------
	*/
	/* Create a new XMLHttpRequest object to talk to the Web server */
	var oBao = false;
	/*@cc_on @*/
	/*@if (@_jscript_version >= 5)
	try {
	  oBao = new ActiveXObject("Msxml2.XMLHTTP");
	} catch (e) {
	  try {
		oBao = new ActiveXObject("Microsoft.XMLHTTP");
	  } catch (e2) {
		oBao = false;
	  }
	}
	@end @*/
	if (!oBao && typeof XMLHttpRequest != 'undefined') {
	  oBao = new XMLHttpRequest();
	}
	//�����ַ���+,%,&,=,?�ȵĴ������취.�ַ�������escape�����.
	//Update:2004-6-1 12:22
		var userInfoo = "username="+username;
		oBao.open("POST","checkuser.asp",false);
		oBao.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		oBao.send(userInfoo);
	//�������˴����ص��Ǿ���escape������ַ���.
	var strResult = unescape(oBao.responseText);
	//alert (strResult);
	//���ַ����ֿ�.
	if (strResult==getObj("input1").value.toLowerCase()){
	getObj("sps1").innerHTML="<img src='img_ext/error.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#ff0000'>����������������Q��</font>";
	mon1=0;
	allok();
	return false;
	}else{
	getObj("sps1").innerHTML="<img src='img_ext/ok.gif' width='16' height='16' hspace='3' align='absmiddle' /><font color='#AED231'>�û���OK�������Խ����²�����</font>";
	mon1=1;
	allok();
	return false;
	}
 }
 }
 }
}

//����Ťť״̬���
function allok(){
  if (mon1==1&&mon2==1&&mon3==1&&mon5==1){
getObj("submit").disabled=false;
}else{
getObj("submit").disabled=true;
}
}
